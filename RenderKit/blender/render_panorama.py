"""Revit から書き出した 3D モデルを Blender で 360°パノラマにレンダリングし、
ブラウザで開ける HTML ファイル1つにまとめる。

Blender の画面は開かず、裏側（バックグラウンド）で動かす:
    blender -b --python-exit-code 1 -P render_panorama.py -- --input モデル.fbx [オプション]

通常は RenderKit/MakePanorama.bat（FBX をドラッグ＆ドロップ）から呼ばれる。

主なオプション:
    --input PATH          読み込むモデル（.fbx / .gltf / .glb / .obj）
    --output PATH         出力する HTML（省略時はモデルと同じフォルダに「<名前>_360.html」）
    --quality NAME        draft（下書き）/ normal（標準）/ high（高品質）
    --eye X,Y,Z           撮影位置（メートル）。複数回指定すると、その数だけパノラマを作る
    --eye-height M        撮影位置を自動で決めるときの目の高さ（既定 1.5m）
    --sun-azimuth DEG     太陽の方位。0=北（モデルの +Y 方向）、90=東。既定 135（南東）
    --sun-elevation DEG   太陽の高さ。0=地平線、90=真上。既定 40
    --exposure EV         明るさの補正。+1 で2倍明るく、-1 で半分（自動露出に上乗せ）
    --no-auto-exposure    自動露出・自動ホワイトバランスを使わない（--exposure の値だけで明るさを決める）
    --device auto|cpu     auto なら GPU が使えれば GPU で計算する
    --save-blend          Blender のファイル（.blend）も保存する（手で微調整したいとき用）

撮影位置の決め方（優先順）:
    1. --eye の指定
    2. モデルに入っているカメラ（Revit の 3D ビュー（パース）から書き出すと入ることがある）
    3. モデル全体の中心の、一番低い位置から目の高さだけ上
"""

import argparse
import math
import os
import sys
import time

import bpy
import numpy as np
from mathutils import Vector

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(SCRIPT_DIR))
import make_viewer  # noqa: E402  (RenderKit/make_viewer.py)

MIN_BLENDER = (4, 2, 0)

# 品質ごとの (画像の横幅ピクセル, 1ピクセルあたりの計算回数)
QUALITY = {
    "draft": (2048, 64),
    "normal": (4096, 256),
    "high": (8192, 512),
}

GLASS_WORDS = ("glass", "glazing", "ガラス", "硝子", "玻璃")
MIRROR_WORDS = ("mirror", "鏡", "镜")
METAL_WORDS = ("metal", "steel", "aluminum", "aluminium", "stainless", "chrome",
               "金属", "スチール", "アルミ", "ステンレス", "鋼", "钢", "铝")

# 空と太陽の明るさ。露出 0 で、晴れた日中の屋外がちょうど良い明るさになるよう調整した値
SKY_STRENGTH = 0.25
SUN_STRENGTH = 4.0
SUN_ANGLE_DEG = 0.53  # 実際の太陽の見かけの大きさ（影のぼけ具合が決まる）

# 自動露出: 画面全体の明るさの平均（対数平均）をこの値にそろえる。写真の「標準の明るさ（18%グレー）」の考え方
AUTO_EXPOSURE_TARGET = 0.18
AUTO_EXPOSURE_RANGE = (-4.0, 6.0)  # 補正しすぎないための上下限（EV）
# 自動ホワイトバランス: 空の光だけだと全体が青っぽくなるので、平均の色の偏りをこの割合だけ打ち消す。
# 1.0 にすると木の床などの本来の色まで抜けてしまうため、半分程度にとどめる
AUTO_WHITE_BALANCE_STRENGTH = 0.5


def log(msg):
    print(f"[RenderKit] {msg}", flush=True)


def parse_args():
    argv = sys.argv[sys.argv.index("--") + 1:] if "--" in sys.argv else []
    p = argparse.ArgumentParser(prog="render_panorama.py")
    p.add_argument("--input", required=True)
    p.add_argument("--output")
    p.add_argument("--quality", choices=QUALITY.keys(), default="normal")
    p.add_argument("--eye", action="append", default=[])
    p.add_argument("--eye-height", type=float, default=1.5)
    p.add_argument("--sun-azimuth", type=float, default=135.0)
    p.add_argument("--sun-elevation", type=float, default=40.0)
    p.add_argument("--exposure", type=float, default=0.0)
    p.add_argument("--no-auto-exposure", action="store_true")
    p.add_argument("--device", choices=("auto", "cpu"), default="auto")
    p.add_argument("--title")
    p.add_argument("--save-blend", action="store_true")
    return p.parse_args(argv)


# --- モデルの読み込み --------------------------------------------------------

def import_model(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".fbx":
        bpy.ops.import_scene.fbx(filepath=path)
    elif ext in (".gltf", ".glb"):
        bpy.ops.import_scene.gltf(filepath=path)
    elif ext == ".obj":
        bpy.ops.wm.obj_import(filepath=path)
    else:
        raise SystemExit(f"対応していない形式です（.fbx / .gltf / .glb / .obj のみ）: {path}")


def scene_bounds(objects):
    lo = Vector((math.inf,) * 3)
    hi = Vector((-math.inf,) * 3)
    for obj in objects:
        for corner in obj.bound_box:
            v = obj.matrix_world @ Vector(corner)
            lo = Vector(map(min, lo, v))
            hi = Vector(map(max, hi, v))
    return lo, hi


# --- 素材（マテリアル）の調整 -------------------------------------------------
# Revit の FBX はマテリアルの情報が少ないため、写真らしく見えるように名前から推測して整える。

def set_input(node, names, value):
    """Blender のバージョンで入力名が違うことがあるので、候補を順に試す"""
    for name in names:
        sock = node.inputs.get(name)
        if sock is not None:
            for link in list(sock.links):
                link.id_data.links.remove(link)
            sock.default_value = value
            return True
    return False


def get_input_value(node, name, default):
    sock = node.inputs.get(name)
    if sock is None or sock.is_linked:
        return default
    return sock.default_value


def remove_missing_textures():
    """画像ファイルが見つからないテクスチャを外す（外さないとピンク色で描かれる）"""
    missing = set()
    for img in bpy.data.images:
        if img.source != "FILE" or img.packed_file:
            continue
        if not os.path.isfile(bpy.path.abspath(img.filepath)):
            missing.add(img.name)
    if not missing:
        return 0
    removed = 0
    for mat in bpy.data.materials:
        if not mat.node_tree:
            continue
        for node in list(mat.node_tree.nodes):
            if node.type == "TEX_IMAGE" and node.image and node.image.name in missing:
                mat.node_tree.nodes.remove(node)
                removed += 1
    log(f"見つからないテクスチャ画像 {len(missing)} 個を外しました（色だけで描きます）")
    return removed


def make_glass(mat, principled):
    tree = mat.node_tree
    set_input(principled, ["Base Color"], (0.9, 0.95, 0.95, 1.0))
    set_input(principled, ["Alpha"], 1.0)
    set_input(principled, ["Roughness"], 0.02)
    set_input(principled, ["IOR"], 1.5)
    set_input(principled, ["Metallic"], 0.0)
    set_input(principled, ["Transmission Weight", "Transmission"], 1.0)

    # ガラスを通った太陽光の計算は時間がかかり、室内が暗くなりがち。
    # 影の計算だけはガラスを透明として扱い、室内に光を入れる（建築パースの定番の手法）。
    output = next((n for n in tree.nodes if n.type == "OUTPUT_MATERIAL" and n.is_active_output), None)
    if output is None:
        return
    light_path = tree.nodes.new("ShaderNodeLightPath")
    transparent = tree.nodes.new("ShaderNodeBsdfTransparent")
    mix = tree.nodes.new("ShaderNodeMixShader")
    tree.links.new(light_path.outputs["Is Shadow Ray"], mix.inputs[0])
    tree.links.new(principled.outputs[0], mix.inputs[1])
    tree.links.new(transparent.outputs[0], mix.inputs[2])
    tree.links.new(mix.outputs[0], output.inputs["Surface"])


def improve_materials():
    counts = {"glass": 0, "mirror": 0, "metal": 0, "other": 0}
    for mat in bpy.data.materials:
        if not mat.node_tree:
            continue
        principled = next((n for n in mat.node_tree.nodes if n.type == "BSDF_PRINCIPLED"), None)
        if principled is None:
            continue
        name = mat.name.lower()
        alpha = get_input_value(principled, "Alpha", 1.0)
        if any(w in name for w in GLASS_WORDS) or alpha < 0.9:
            make_glass(mat, principled)
            counts["glass"] += 1
        elif any(w in name for w in MIRROR_WORDS):
            set_input(principled, ["Metallic"], 1.0)
            set_input(principled, ["Roughness"], 0.02)
            set_input(principled, ["Base Color"], (0.95, 0.95, 0.95, 1.0))
            counts["mirror"] += 1
        elif any(w in name for w in METAL_WORDS):
            set_input(principled, ["Metallic"], 1.0)
            rough = get_input_value(principled, "Roughness", 0.3)
            set_input(principled, ["Roughness"], min(max(rough, 0.2), 0.45))
            counts["metal"] += 1
        else:
            # FBX の「光沢」から変換された値はテカテカになりやすいので、ほどよいざらつきにそろえる
            rough = get_input_value(principled, "Roughness", 0.5)
            set_input(principled, ["Roughness"], min(max(rough, 0.4), 0.9))
            set_input(principled, ["Metallic"], 0.0)
            counts["other"] += 1
    log(f"マテリアル調整: ガラス {counts['glass']} / 鏡 {counts['mirror']} / 金属 {counts['metal']} / その他 {counts['other']}")


# --- 空と太陽 ----------------------------------------------------------------

def sun_vector(azimuth_deg, elevation_deg):
    """太陽のある方向（北=+Y、東=+X、上=+Z）"""
    az, el = math.radians(azimuth_deg), math.radians(elevation_deg)
    return Vector((math.sin(az) * math.cos(el), math.cos(az) * math.cos(el), math.sin(el)))


def setup_lighting(scene, azimuth, elevation):
    world = bpy.data.worlds.new("RenderKit_Sky")
    scene.world = world
    if world.node_tree is None:  # Blender 4.x では明示的にノードを有効にする必要がある
        world.use_nodes = True
    tree = world.node_tree
    tree.nodes.clear()
    sky = tree.nodes.new("ShaderNodeTexSky")
    # Blender 5.0 で「NISHITA」が「MULTIPLE_SCATTERING」に変わったため、使える方を選ぶ
    sky_types = sky.bl_rna.properties["sky_type"].enum_items.keys()
    for t in ("MULTIPLE_SCATTERING", "NISHITA", "HOSEK_WILKIE"):
        if t in sky_types:
            sky.sky_type = t
            break
    if hasattr(sky, "sun_disc"):
        sky.sun_disc = False  # 太陽そのものは下の「太陽光」で表現する（ノイズが少ない）
    if hasattr(sky, "sun_elevation"):
        sky.sun_elevation = math.radians(max(elevation, 1.0))
        # 空の太陽の向きは、+Y から時計回りの角度で指定する
        sky.sun_rotation = math.radians(azimuth)
    if sky.sky_type == "HOSEK_WILKIE":
        sky.sun_direction = sun_vector(azimuth, elevation)
    bg = tree.nodes.new("ShaderNodeBackground")
    bg.inputs["Strength"].default_value = SKY_STRENGTH
    out = tree.nodes.new("ShaderNodeOutputWorld")
    tree.links.new(sky.outputs["Color"], bg.inputs["Color"])
    tree.links.new(bg.outputs["Background"], out.inputs["Surface"])

    sun_data = bpy.data.lights.new("RenderKit_Sun", type="SUN")
    sun_data.energy = SUN_STRENGTH if elevation > 0 else 0.0
    sun_data.angle = math.radians(SUN_ANGLE_DEG)
    sun = bpy.data.objects.new("RenderKit_Sun", sun_data)
    scene.collection.objects.link(sun)
    # 太陽光はオブジェクトの -Z 方向へ進むので、+Z を太陽の方へ向ける
    sun.rotation_euler = sun_vector(azimuth, elevation).to_track_quat("Z", "Y").to_euler()


# --- 撮影位置 ----------------------------------------------------------------

def parse_eyes(text):
    """「X,Y,Z」を位置のリストにする。数字を3つずつ区切って読むので、
    「X,Y,Z;X,Y,Z」や、PowerShell 経由でつながった「X,Y,Z,X,Y,Z」でも複数の位置として扱える"""
    try:
        numbers = [float(v) for v in text.replace(";", ",").replace(" ", ",").split(",") if v]
    except ValueError:
        numbers = []
    if not numbers or len(numbers) % 3 != 0:
        raise SystemExit(f"--eye は「X,Y,Z」（複数なら「X,Y,Z;X,Y,Z」）の形で指定してください: {text}")
    return [Vector(numbers[i:i + 3]) for i in range(0, len(numbers), 3)]


def collect_viewpoints(args, imported_cameras, bounds):
    """[(名前, 位置, 向き(度), 見上げ角(度))] を返す"""
    if args.eye:
        eyes = [eye for text in args.eye for eye in parse_eyes(text)]
        return [(f"View {i + 1}", eye, 0.0, 0.0) for i, eye in enumerate(eyes)]

    points = []
    for cam in imported_cameras:
        m = cam.matrix_world
        forward = m.to_3x3() @ Vector((0.0, 0.0, -1.0))
        heading = math.degrees(math.atan2(-forward.x, forward.y))
        pitch = math.degrees(math.asin(max(-1.0, min(1.0, forward.normalized().z))))
        points.append((cam.name, m.translation.copy(), heading, pitch))
    if points:
        log(f"モデルに入っていたカメラ {len(points)} 個の位置から撮影します")
        return points

    lo, hi = bounds
    center = (lo + hi) / 2
    eye = Vector((center.x, center.y, lo.z + args.eye_height))
    log("カメラが無いため、モデルの中心から撮影します（位置は --eye で指定できます）")
    return [("View 1", eye, 0.0, 0.0)]


def create_panorama_camera(scene):
    cam_data = bpy.data.cameras.new("RenderKit_Panorama")
    cam_data.type = "PANO"
    cam_data.panorama_type = "EQUIRECTANGULAR"
    cam_data.clip_start = 0.05
    cam_data.clip_end = 100000.0
    cam = bpy.data.objects.new("RenderKit_Panorama", cam_data)
    scene.collection.objects.link(cam)
    scene.camera = cam
    return cam


# --- レンダリングの設定 ------------------------------------------------------

def setup_device(scene, mode):
    scene.cycles.device = "CPU"
    if mode == "cpu":
        log("計算装置: CPU（指定）")
        return
    try:
        prefs = bpy.context.preferences.addons["cycles"].preferences
    except KeyError:
        log("計算装置: CPU（GPU の設定を読めませんでした）")
        return
    for device_type in ("OPTIX", "CUDA", "HIP", "ONEAPI", "METAL"):
        try:
            prefs.compute_device_type = device_type
        except TypeError:
            continue  # この PC・この Blender では使えない種類
        prefs.get_devices()
        gpus = [d for d in prefs.devices if d.type == device_type]
        if gpus:
            for d in prefs.devices:
                d.use = d.type == device_type
            scene.cycles.device = "GPU"
            log(f"計算装置: GPU（{device_type}: {', '.join(d.name for d in gpus)}）")
            return
    log("計算装置: CPU（使える GPU が見つかりませんでした。時間がかかります）")


def setup_render(scene, quality):
    width, samples = QUALITY[quality]
    scene.render.engine = "CYCLES"
    scene.render.resolution_x = width
    scene.render.resolution_y = width // 2
    scene.render.resolution_percentage = 100

    c = scene.cycles
    c.samples = samples
    c.use_adaptive_sampling = True
    c.adaptive_threshold = 0.01
    c.use_denoising = True
    try:
        c.denoiser = "OPENIMAGEDENOISE"  # CPU でも GPU でも使えるノイズ除去
    except TypeError:
        pass
    c.max_bounces = 8
    c.diffuse_bounces = 4
    c.glossy_bounces = 4
    c.transmission_bounces = 8
    c.transparent_max_bounces = 16
    c.caustics_reflective = False
    c.caustics_refractive = False
    # 強い光の粒（ホタルのような白い点）を抑える
    c.sample_clamp_indirect = 10.0

    # 色の仕上げ（明るい窓の外も白飛びしにくい AgX に、少しメリハリを付ける）。
    # 選べる値の一覧は実行時にしか分からないため、設定してみて失敗したら次の候補を試す。
    for view_transform, look in (("AgX", "AgX - Medium High Contrast"), ("Filmic", "Filmic - Medium High Contrast")):
        try:
            scene.view_settings.view_transform = view_transform
        except TypeError:
            continue
        try:
            scene.view_settings.look = look
        except TypeError:
            pass
        break

    img = scene.render.image_settings
    img.file_format = "JPEG"
    img.quality = 90
    log(f"品質: {quality}（{width}×{width // 2} ピクセル、最大 {samples} サンプル）")


def measure_scene(scene, tmp_path):
    """小さく軽くレンダリングして全体の明るさと色の偏りを測り、
    (ちょうど良く見える露出 EV, 白とみなす色) を返す。
    カメラの自動露出・自動ホワイトバランスと同じ考え方で、暗い室内は明るく、まぶしい屋外は抑える。"""
    r, c, img_settings = scene.render, scene.cycles, scene.render.image_settings
    saved = (r.resolution_x, r.resolution_y, c.samples, img_settings.file_format, r.filepath)
    r.resolution_x, r.resolution_y = 512, 256
    c.samples = 16
    img_settings.file_format = "OPEN_EXR"  # 色の補正がかかる前の、本来の明るさのまま保存される
    r.filepath = tmp_path
    try:
        bpy.ops.render.render(write_still=True)
        img = bpy.data.images.load(tmp_path)
        pixels = np.empty(len(img.pixels), dtype=np.float32)
        img.pixels.foreach_get(pixels)
        bpy.data.images.remove(img)
    finally:
        r.resolution_x, r.resolution_y, c.samples, img_settings.file_format, r.filepath = saved
        if os.path.isfile(tmp_path):
            os.remove(tmp_path)

    rgb = pixels.reshape(-1, 4)[:, :3]
    luminance = rgb @ np.array([0.2126, 0.7152, 0.0722], dtype=np.float32)
    log_average = float(np.exp(np.mean(np.log(np.maximum(luminance, 1e-4)))))
    ev = math.log2(AUTO_EXPOSURE_TARGET / log_average)
    ev = min(max(ev, AUTO_EXPOSURE_RANGE[0]), AUTO_EXPOSURE_RANGE[1])

    # 色の偏りは、窓の外など明るすぎる部分（上位 5%）を除いた平均の色で測る
    usable = luminance <= np.percentile(luminance, 95)
    mean_rgb = rgb[usable].mean(axis=0) if usable.any() else rgb.mean(axis=0)
    mean_rgb = mean_rgb / max(float(mean_rgb @ np.array([0.2126, 0.7152, 0.0722])), 1e-6)
    whitepoint = 1.0 + (mean_rgb - 1.0) * AUTO_WHITE_BALANCE_STRENGTH
    return ev, tuple(float(v) for v in whitepoint)


# --- メイン ------------------------------------------------------------------

def main():
    if bpy.app.version < MIN_BLENDER:
        raise SystemExit(f"Blender {'.'.join(map(str, MIN_BLENDER))} 以降が必要です（今は {bpy.app.version_string}）")
    args = parse_args()
    model = os.path.abspath(args.input)
    if not os.path.isfile(model):
        raise SystemExit(f"モデルが見つかりません: {model}")

    base = os.path.splitext(os.path.basename(model))[0]
    out_html = os.path.abspath(args.output) if args.output else os.path.join(os.path.dirname(model), f"{base}_360.html")
    out_dir = os.path.dirname(out_html)
    out_stem = os.path.splitext(os.path.basename(out_html))[0]

    started = time.time()
    log(f"Blender {bpy.app.version_string}")
    bpy.ops.wm.read_factory_settings(use_empty=True)
    scene = bpy.context.scene

    log(f"読み込み中: {model}")
    import_model(model)
    meshes = [o for o in scene.objects if o.type == "MESH"]
    cameras = [o for o in scene.objects if o.type == "CAMERA"]
    if not meshes:
        raise SystemExit("モデルに形状がありません。Revit で 3D ビューを開いた状態で書き出したか確認してください")
    bounds = scene_bounds(meshes)
    size = bounds[1] - bounds[0]
    log(f"形状 {len(meshes)} 個、全体の大きさ {size.x:.1f} × {size.y:.1f} × {size.z:.1f} m")
    if max(size) > 5000 or max(size) < 0.5:
        log("⚠ 大きさが建物として不自然です。書き出し時の単位設定を確認してください")

    remove_missing_textures()
    improve_materials()
    setup_lighting(scene, args.sun_azimuth, args.sun_elevation)
    setup_render(scene, args.quality)
    setup_device(scene, args.device)

    viewpoints = collect_viewpoints(args, cameras, bounds)
    pano = create_panorama_camera(scene)

    panoramas = []
    for i, (name, eye, heading, pitch) in enumerate(viewpoints, start=1):
        image_path = os.path.join(out_dir, f"{out_stem}_{i}.jpg")
        pano.location = eye
        # 画像の中央が heading の方向（0=北）になるように水平に向ける
        pano.rotation_euler = (math.radians(90.0), 0.0, math.radians(heading))
        log(f"レンダリング {i}/{len(viewpoints)}: {name}  位置 ({eye.x:.2f}, {eye.y:.2f}, {eye.z:.2f})")
        t0 = time.time()
        exposure = args.exposure
        view = scene.view_settings
        if not args.no_auto_exposure:
            if hasattr(view, "use_white_balance"):
                view.use_white_balance = False  # 測るときは補正なしの色で測る
            auto_ev, whitepoint = measure_scene(scene, os.path.join(out_dir, f"{out_stem}_{i}_measure.exr"))
            exposure += auto_ev
            log(f"  自動露出 {auto_ev:+.1f} EV（補正 {args.exposure:+.1f} → 合計 {exposure:+.1f} EV）")
            if hasattr(view, "white_balance_whitepoint"):  # Blender 4.3 以降
                view.use_white_balance = True
                view.white_balance_whitepoint = whitepoint
                log(f"  自動ホワイトバランス {view.white_balance_temperature:.0f} K")
        view.exposure = exposure
        scene.render.filepath = image_path
        bpy.ops.render.render(write_still=True)
        log(f"  完了（{time.time() - t0:.0f} 秒）→ {image_path}")
        panoramas.append({"name": name, "path": image_path, "yaw": 0.0, "pitch": pitch})

    if args.save_blend:
        blend_path = os.path.join(out_dir, f"{out_stem}.blend")
        bpy.ops.wm.save_as_mainfile(filepath=blend_path)
        log(f"Blender ファイルを保存しました: {blend_path}")

    make_viewer.build_html(panoramas, args.title or base, out_html)
    size_mb = os.path.getsize(out_html) / (1024 * 1024)
    log(f"完成: {out_html}（{size_mb:.1f} MB、合計 {time.time() - started:.0f} 秒）")


if __name__ == "__main__":
    main()

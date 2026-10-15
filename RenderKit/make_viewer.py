"""360°パノラマ画像を、ブラウザで開ける HTML ファイル1つにまとめる。

画像は HTML の中に埋め込む（Base64 という文字の形に変換して入れる）ので、
出来た HTML だけを渡せば、どの PC・スマホでもネット接続なしで見られる。

使い方:
    python make_viewer.py 出力.html 画像1.jpg [画像2.jpg ...] [--title タイトル]

画像の名前（切り替えボタンの表示名）はファイル名から作る。
Blender のスクリプト（blender/render_panorama.py）からも build_html() を呼び出して使う。
"""

import argparse
import base64
import html
import json
import os
import sys

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "viewer", "panorama_viewer.html")
PANORAMA_MARKER = "/*__PANORAMAS__*/[]"
TITLE_MARKER = "__TITLE__"

MIME_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
}


def image_to_data_uri(path):
    ext = os.path.splitext(path)[1].lower()
    mime = MIME_TYPES.get(ext)
    if mime is None:
        raise ValueError(f"対応していない画像形式です（jpg / png / webp のみ）: {path}")
    with open(path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def build_html(panoramas, title, out_path):
    """panoramas: [{"name": 表示名, "path": 画像ファイル, "yaw": 最初に向く方向(度・省略可)}]"""
    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        template = f.read()
    if PANORAMA_MARKER not in template or TITLE_MARKER not in template:
        raise RuntimeError(f"ビューアのひな形に埋め込み位置の目印がありません: {TEMPLATE_PATH}")

    data = [
        {
            "name": p.get("name") or os.path.splitext(os.path.basename(p["path"]))[0],
            "yaw": float(p.get("yaw", 0.0)),
            "pitch": float(p.get("pitch", 0.0)),
            "src": image_to_data_uri(p["path"]),
        }
        for p in panoramas
    ]
    # 名前に "</script>" が入っていてもスクリプトが途中で切れないようにする
    data_js = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")

    page = template.replace(PANORAMA_MARKER, data_js, 1).replace(TITLE_MARKER, html.escape(title), 1)
    out_dir = os.path.dirname(os.path.abspath(out_path))
    os.makedirs(out_dir, exist_ok=True)
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(page)
    return out_path


def main(argv=None):
    parser = argparse.ArgumentParser(description="360°パノラマ画像を HTML ファイル1つにまとめる")
    parser.add_argument("output", help="出力する HTML ファイル")
    parser.add_argument("images", nargs="+", help="正距円筒図法（横:縦 = 2:1）のパノラマ画像")
    parser.add_argument("--title", default="360° Panorama", help="ページのタイトル")
    args = parser.parse_args(argv)

    for path in args.images:
        if not os.path.isfile(path):
            parser.error(f"画像が見つかりません: {path}")

    out = build_html([{"path": p} for p in args.images], args.title, args.output)
    size_mb = os.path.getsize(out) / (1024 * 1024)
    print(f"作成しました: {out}（{size_mb:.1f} MB）")
    return 0


if __name__ == "__main__":
    sys.exit(main())

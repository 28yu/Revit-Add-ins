#include "i18n.h"

#include <cstring>

namespace t28 {

static Lang g_lang = Lang::JP;

struct Entry {
    const char* key;
    const char* jp;
    const char* us;
    const char* cn;
};

static const Entry kStrings[] = {
    {"Crashed",
     "描画プログラムが予期せず終了しました。\n"
     "記録（%LOCALAPPDATA%\\Tools28\\Render\\render.log）を開発者に送っていただけると、原因を調べられます。",
     "The renderer stopped unexpectedly.\n"
     "Please send the log (%LOCALAPPDATA%\\Tools28\\Render\\render.log) to the developer so the cause can be investigated.",
     "渲染程序意外终止。\n"
     "请将日志（%LOCALAPPDATA%\\Tools28\\Render\\render.log）发送给开发者，以便调查原因。"},
    {"CrashedStack",
     "描画プログラムが予期せず終了しました（処理の深さの上限を超えました）。\n"
     "記録（%LOCALAPPDATA%\\Tools28\\Render\\render.log）を開発者に送ってください。",
     "The renderer stopped unexpectedly (stack overflow).\n"
     "Please send the log (%LOCALAPPDATA%\\Tools28\\Render\\render.log) to the developer.",
     "渲染程序意外终止（超出处理深度上限）。\n"
     "请将日志（%LOCALAPPDATA%\\Tools28\\Render\\render.log）发送给开发者。"},
    {"OutOfMemory",
     "メモリが足りないため、描画を続けられませんでした。\n"
     "ほかのアプリを閉じるか、3D ビューの切断ボックスで表示する範囲を狭めてから、もう一度お試しください。",
     "Not enough memory to continue rendering.\n"
     "Close other apps or narrow the 3D view with a section box, then try again.",
     "内存不足，无法继续渲染。\n"
     "请关闭其他应用程序，或用剖面框缩小 3D 视图的显示范围后重试。"},
    {"OutOfGpuMemory",
     "GPU のメモリが足りないため、描画を続けられませんでした。\n"
     "3D ビューの切断ボックスで表示する範囲を狭めるか、表示するカテゴリを減らしてから、もう一度お試しください。",
     "Not enough GPU memory to continue rendering.\n"
     "Narrow the 3D view with a section box or hide some categories, then try again.",
     "GPU 内存不足，无法继续渲染。\n"
     "请用剖面框缩小 3D 视图的显示范围或减少显示的类别后重试。"},
    {"DeviceLost",
     "GPU の処理に時間がかかりすぎたため、Windows が GPU をリセットしました。\n"
     "3D ビューの切断ボックスで表示する範囲を狭めるか、表示するカテゴリを減らしてから、もう一度お試しください。",
     "The GPU took too long and Windows reset it.\n"
     "Narrow the 3D view with a section box or hide some categories, then try again.",
     "GPU 处理时间过长，Windows 已重置 GPU。\n"
     "请用剖面框缩小 3D 视图的显示范围或减少显示的类别后重试。"},
    {"ModelTooLarge",
     "モデルが大きすぎて、この GPU では扱えません。\n"
     "3D ビューの切断ボックスで表示する範囲を狭めるか、表示するカテゴリを減らしてから、もう一度お試しください。",
     "The model is too large for this GPU.\n"
     "Narrow the 3D view with a section box or hide some categories, then try again.",
     "模型过大，此 GPU 无法处理。\n"
     "请用剖面框缩小 3D 视图的显示范围或减少显示的类别后重试。"},
    {"QualityLowered",
     "モデルが大きいため、解像度を下げて描いています",
     "The model is large, so it is drawn at a lower resolution",
     "模型较大，正在以较低分辨率渲染"},
    {"NoVulkan",
     "Vulkan（GPU を動かす仕組み）を使えません。GPU のドライバーを最新にしてください。",
     "Vulkan is not available. Please update your GPU driver.",
     "无法使用 Vulkan。请将显卡驱动更新到最新版本。"},
    {"NoGpu",
     "使える GPU が見つかりません。GPU のドライバーを最新にしてください。\n"
     "（リモートデスクトップや仮想マシンでは GPU を使えないことがあります）",
     "No usable GPU was found. Please update your GPU driver.\n"
     "(The GPU may not be available over Remote Desktop or in a virtual machine.)",
     "未找到可用的 GPU。请将显卡驱动更新到最新版本。\n"
     "（通过远程桌面或在虚拟机中可能无法使用 GPU）"},
    {"SceneLoadFailed",
     "モデルのデータを読み込めませんでした。",
     "Could not load the model data.",
     "无法读取模型数据。"},
    {"WindowTitle",
     "Tools28 リアルタイムレンダリング",
     "Tools28 Real-time Rendering",
     "Tools28 实时渲染"},
    // 操作の案内（ウィンドウの左下に小さく出す）。絵の下に付ける短い言葉
    {"HelpZoom", "ズーム", "Zoom", "缩放"},
    {"HelpPan", "移動", "Pan", "平移"},
    {"HelpOrbit", "回転", "Orbit", "动态观察"},
    {"HelpFit", "全体表示", "Zoom to fit", "缩放匹配"},
    {"HelpLook", "見回す", "Look around", "环视"},
    {"HelpWalk", "歩く", "Walk", "行走"},
    {"HelpUpDown", "下がる・上がる", "Down / up", "下降・上升"},
    {"HelpBright", "明るさ", "Brightness", "亮度"},
    {"HelpReset", "最初の視点", "Reset view", "初始视点"},
    // 書き出し（右上のボタン）
    {"ExportImage", "画像", "Image", "图像"},
    {"ExportPanorama", "360°", "360°", "360°"},
    {"ExportViewer", "ビューア", "Viewer", "查看器"},
    {"TipImage", "今の視点を高画質の画像（PNG）で保存します（P キー）",
     "Save the current view as a high-quality image (PNG) (P key)",
     "将当前视点保存为高画质图像（PNG）（P 键）"},
    {"TipPanorama", "この位置から見回せる 360°画像を HTML ファイルで保存します（ブラウザで開けます）",
     "Save a 360° image you can look around in as an HTML file (opens in a web browser)",
     "将可环视的 360° 图像保存为 HTML 文件（可在浏览器中打开）"},
    {"TipViewer", "Revit の無い PC でも開ける、歩き回れるビューアを exe ファイルで保存します",
     "Save a walk-through viewer as an exe file that opens on PCs without Revit",
     "将可漫游的查看器保存为 exe 文件，无需 Revit 即可打开"},
    {"ExportingImage", "画像を書き出し中…", "Exporting image...", "正在导出图像…"},
    {"ExportingPanorama", "360°画像を書き出し中…", "Exporting 360° image...", "正在导出 360° 图像…"},
    {"ExportCancelHint", "Esc キーで中止", "Press Esc to cancel", "按 Esc 键取消"},
    {"Saved", "保存しました: ", "Saved: ", "已保存: "},
    {"SavedPanoramaNote", "ブラウザで開くと、ドラッグで見回せます", "Open it in a web browser and drag to look around",
     "在浏览器中打开后，拖动即可环视"},
    {"SavedViewerNote", "ダブルクリックで開けます（Revit は不要。Windows の警告が出たら「詳細情報」→「実行」）",
     "Double-click to open (no Revit needed. If Windows warns you, click \"More info\" > \"Run anyway\")",
     "双击即可打开（无需 Revit。如出现 Windows 警告，请点击“更多信息”→“仍要运行”）"},
    {"ExportFailed", "保存できませんでした: ", "Could not save: ", "无法保存: "},
    {"ExportCancelled", "書き出しを中止しました", "Export cancelled", "已取消导出"},
    // ライブ同期（Revit での変更を受け取る）
    {"SyncOn", "同期中", "Sync on", "同步中"},
    {"SyncOff", "同期停止", "Sync off", "已停止同步"},
    {"TipSync", "Revit でのモデルの変更を自動で反映します（視点は連動しません。クリックで止める / 再開）",
     "Automatically reflects model changes made in Revit (the view is not linked; click to pause / resume)",
     "自动反映 Revit 中的模型修改（视点不联动；单击可暂停 / 恢复）"},
    {"SyncUpdated", "Revit の変更を反映しました", "Updated with the changes in Revit", "已反映 Revit 中的修改"},
    {"SyncPaused", "同期を止めました（Revit での変更は反映されません）", "Sync paused (changes in Revit are not reflected)",
     "已暂停同步（不会反映 Revit 中的修改）"},
    {"SyncResumed", "同期を再開しました", "Sync resumed", "已恢复同步"},
    {"FilterPng", "PNG 画像 (*.png)", "PNG image (*.png)", "PNG 图像 (*.png)"},
    {"FilterHtml", "360°画像 (*.html)", "360° image (*.html)", "360° 图像 (*.html)"},
    {"FilterExe", "ビューア (*.exe)", "Viewer (*.exe)", "查看器 (*.exe)"},
    {"PanoramaHint", "ドラッグで見回す・ホイールでズーム・ダブルクリックで全画面",
     "Drag to look around / wheel to zoom / double-click for full screen",
     "拖动环视・滚轮缩放・双击全屏"},
    {"PanoramaNoWebGL", "このブラウザでは表示できません（WebGL が使えません）。別のブラウザで開いてください。",
     "This browser cannot show the image (WebGL is not available). Please open it in another browser.",
     "此浏览器无法显示（无法使用 WebGL）。请使用其他浏览器打开。"},
};

void setLanguage(const char* code) {
    if (!code) return;
    if (strcmp(code, "US") == 0 || strcmp(code, "EN") == 0) g_lang = Lang::US;
    else if (strcmp(code, "CN") == 0) g_lang = Lang::CN;
    else g_lang = Lang::JP;
}

Lang currentLanguage() { return g_lang; }

const char* tr(const char* key) {
    for (const Entry& e : kStrings) {
        if (strcmp(e.key, key) == 0) {
            switch (g_lang) {
                case Lang::US: return e.us;
                case Lang::CN: return e.cn;
                default: return e.jp;
            }
        }
    }
    return key;
}

}  // namespace t28

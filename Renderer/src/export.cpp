#include "export.h"

#include <cctype>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <ctime>

#include "i18n.h"
#include "log.h"
#include "scene_pack.h"

#include <stb_image_write.h>

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commdlg.h>
#else
#include <unistd.h>
#endif

namespace t28 {

namespace {

const char kTrailerMagic[8] = {'T', '2', '8', 'E', 'M', 'B', 'E', 'D'};
const char kPayloadMagic[4] = {'T', '2', '8', 'V'};

#ifdef _WIN32
std::wstring widen(const std::string& s) {
    int n = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, nullptr, 0);
    std::wstring w(n > 0 ? n - 1 : 0, L'\0');
    if (n > 1) MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, &w[0], n);
    return w;
}
std::string narrow(const std::wstring& w) {
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string s(n > 0 ? n - 1 : 0, '\0');
    if (n > 1) WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &s[0], n, nullptr, nullptr);
    return s;
}
#endif

std::string selfExePath() {
#ifdef _WIN32
    std::wstring buf(32768, L'\0');
    DWORD n = GetModuleFileNameW(nullptr, &buf[0], static_cast<DWORD>(buf.size()));
    buf.resize(n);
    return narrow(buf);
#else
    char buf[4096];
    ssize_t n = readlink("/proc/self/exe", buf, sizeof(buf) - 1);
    if (n <= 0) return "";
    buf[n] = '\0';
    return buf;
#endif
}

bool readAll(const std::string& path, std::vector<uint8_t>& data) {
    FILE* f = openFile(path, "rb");
    if (!f) return false;
    fseek(f, 0, SEEK_END);
    long size = ftell(f);
    fseek(f, 0, SEEK_SET);
    data.resize(size > 0 ? static_cast<size_t>(size) : 0);
    bool ok = size > 0 && fread(data.data(), 1, data.size(), f) == data.size();
    fclose(f);
    return ok;
}

bool writeAll(const std::string& path, const std::vector<const std::vector<uint8_t>*>& parts) {
    FILE* f = openFile(path, "wb");
    if (!f) return false;
    bool ok = true;
    for (const auto* p : parts) ok = ok && (p->empty() || fwrite(p->data(), 1, p->size(), f) == p->size());
    ok = (fclose(f) == 0) && ok;
    return ok;
}

// exe の後ろに付け足したもの（シーン）が始まる位置。付け足しが無ければ -1
long payloadOffset(const std::vector<uint8_t>& exe) {
    if (exe.size() < 16 || memcmp(&exe[exe.size() - 8], kTrailerMagic, 8) != 0) return -1;
    uint64_t offset = 0;
    memcpy(&offset, &exe[exe.size() - 16], 8);
    return offset < exe.size() - 16 ? static_cast<long>(offset) : -1;
}

std::string base64(const std::vector<uint8_t>& in) {
    static const char* t = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    std::string out;
    out.reserve((in.size() + 2) / 3 * 4);
    for (size_t i = 0; i < in.size(); i += 3) {
        uint32_t v = in[i] << 16;
        if (i + 1 < in.size()) v |= in[i + 1] << 8;
        if (i + 2 < in.size()) v |= in[i + 2];
        out += t[(v >> 18) & 63];
        out += t[(v >> 12) & 63];
        out += i + 1 < in.size() ? t[(v >> 6) & 63] : '=';
        out += i + 2 < in.size() ? t[v & 63] : '=';
    }
    return out;
}

// HTML の中に文字として入れるとき、意味を持つ記号を置き換える
std::string htmlEscape(const std::string& s) {
    std::string out;
    for (char c : s) {
        switch (c) {
            case '&': out += "&amp;"; break;
            case '<': out += "&lt;"; break;
            case '>': out += "&gt;"; break;
            case '"': out += "&quot;"; break;
            default: out += c;
        }
    }
    return out;
}

std::string replaceAll(std::string s, const std::string& from, const std::string& to) {
    for (size_t pos = 0; (pos = s.find(from, pos)) != std::string::npos; pos += to.size()) s.replace(pos, from.size(), to);
    return s;
}

// 360°パノラマを見る HTML（WebGL。外部のファイルやインターネットは使わない）
const char* kPanoramaHtml = R"HTML(<!DOCTYPE html>
<html lang="{{LANG}}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{TITLE}}</title>
<style>
html, body { margin: 0; height: 100%; background: #111; overflow: hidden; font-family: system-ui, sans-serif; }
canvas { display: block; width: 100%; height: 100%; cursor: grab; touch-action: none; }
canvas:active { cursor: grabbing; }
.tag { position: fixed; left: 12px; color: #eee; background: rgba(0, 0, 0, .55); padding: 6px 10px; border-radius: 6px; }
#title { top: 12px; font-size: 14px; }
#hint { bottom: 12px; font-size: 13px; }
#error { position: fixed; inset: 0; display: none; align-items: center; justify-content: center; color: #eee; padding: 24px; text-align: center; }
#save { position: fixed; right: 12px; bottom: 12px; font-size: 13px; color: #eee; background: rgba(0, 0, 0, .55); padding: 6px 10px;
        border-radius: 6px; text-decoration: none; }
#save:hover { background: rgba(0, 0, 0, .75); }
</style>
</head>
<body>
<canvas id="view"></canvas>
<div id="title" class="tag">{{TITLE}}</div>
<div id="hint" class="tag">{{HINT}}</div>
<div id="error">{{NOWEBGL}}</div>
<a id="save" download="{{FILENAME}}">{{SAVE}}</a>
<script>
(function () {
  var canvas = document.getElementById("view");
  var gl = canvas.getContext("webgl");
  if (!gl) { document.getElementById("error").style.display = "flex"; return; }
  var vs = "attribute vec2 p; varying vec2 v; void main() { v = p; gl_Position = vec4(p, 0.0, 1.0); }";
  var fs = [
    "precision highp float;",
    "uniform sampler2D tex; uniform vec2 res; uniform float yaw, pitch, tanHalf; varying vec2 v;",
    "void main() {",
    "  vec3 f = vec3(sin(yaw) * cos(pitch), cos(yaw) * cos(pitch), sin(pitch));",
    "  vec3 r = vec3(cos(yaw), -sin(yaw), 0.0);",
    "  vec3 u = cross(r, f);",
    "  vec3 d = normalize(f + r * v.x * tanHalf * res.x / res.y + u * v.y * tanHalf);",
    "  float lon = atan(d.x, d.y), lat = asin(clamp(d.z, -1.0, 1.0));",
    "  gl_FragColor = texture2D(tex, vec2(fract(lon / 6.2831853 + 0.5), 0.5 - lat / 3.1415927));",
    "}"].join("\n");
  function shader(type, src) { var s = gl.createShader(type); gl.shaderSource(s, src); gl.compileShader(s); return s; }
  var prog = gl.createProgram();
  gl.attachShader(prog, shader(gl.VERTEX_SHADER, vs));
  gl.attachShader(prog, shader(gl.FRAGMENT_SHADER, fs));
  gl.linkProgram(prog);
  gl.useProgram(prog);
  gl.bindBuffer(gl.ARRAY_BUFFER, gl.createBuffer());
  gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 1, -1, -1, 1, 1, 1]), gl.STATIC_DRAW);
  var loc = gl.getAttribLocation(prog, "p");
  gl.enableVertexAttribArray(loc);
  gl.vertexAttribPointer(loc, 2, gl.FLOAT, false, 0, 0);
  var yaw = 0, pitch = 0, fov = 75, ready = false;
  var tex = gl.createTexture();
  var img = new Image();
  img.onload = function () {
    gl.bindTexture(gl.TEXTURE_2D, tex);
    gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGB, gl.RGB, gl.UNSIGNED_BYTE, img);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
    gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
    ready = true;
    draw();
  };
  img.src = "data:image/jpeg;base64,{{IMAGE}}";
  document.getElementById("save").href = img.src;  // 他の 360°ビューアや SNS で使えるよう、中の JPEG を保存できる
  function draw() {
    var dpr = window.devicePixelRatio || 1;
    var w = Math.round(canvas.clientWidth * dpr), h = Math.round(canvas.clientHeight * dpr);
    if (canvas.width !== w || canvas.height !== h) { canvas.width = w; canvas.height = h; }
    gl.viewport(0, 0, w, h);
    if (!ready) return;
    gl.uniform2f(gl.getUniformLocation(prog, "res"), w, h);
    gl.uniform1f(gl.getUniformLocation(prog, "yaw"), yaw);
    gl.uniform1f(gl.getUniformLocation(prog, "pitch"), pitch);
    gl.uniform1f(gl.getUniformLocation(prog, "tanHalf"), Math.tan(fov * Math.PI / 360));
    gl.drawArrays(gl.TRIANGLE_STRIP, 0, 4);
  }
  var dragging = false, lx = 0, ly = 0;
  canvas.addEventListener("pointerdown", function (e) { dragging = true; lx = e.clientX; ly = e.clientY; canvas.setPointerCapture(e.pointerId); });
  canvas.addEventListener("pointerup", function () { dragging = false; });
  canvas.addEventListener("pointermove", function (e) {
    if (!dragging) return;
    var k = 2 * Math.tan(fov * Math.PI / 360) / canvas.clientHeight;
    yaw -= (e.clientX - lx) * k;
    pitch = Math.max(-1.55, Math.min(1.55, pitch + (e.clientY - ly) * k));
    lx = e.clientX; ly = e.clientY;
    draw();
  });
  canvas.addEventListener("wheel", function (e) {
    e.preventDefault();
    fov = Math.max(30, Math.min(100, fov * (e.deltaY > 0 ? 1.08 : 1 / 1.08)));
    draw();
  }, { passive: false });
  canvas.addEventListener("dblclick", function () {
    if (document.fullscreenElement) document.exitFullscreen(); else document.documentElement.requestFullscreen();
  });
  window.addEventListener("resize", draw);
  draw();
})();
</script>
</body>
</html>
)HTML";

}  // namespace

std::string fileNameOf(const std::string& path) {
    size_t p = path.find_last_of("/\\");
    return p == std::string::npos ? path : path.substr(p + 1);
}

std::string defaultExportName(const std::string& viewName, ExportKind kind, bool png) {
    std::string base;
    for (unsigned char c : viewName) {
        if (c < 0x20 || strchr("\\/:*?\"<>|", c)) base += '_';
        else base += static_cast<char>(c);
    }
    // 長すぎる題名は短くする（UTF-8 の文字の途中で切らない）
    if (base.size() > 60) {
        size_t cut = 60;
        while (cut > 0 && (static_cast<unsigned char>(base[cut]) & 0xC0) == 0x80) --cut;
        base.resize(cut);
    }
    if (base.empty()) base = "Tools28";
    // 日付（年の下 2 桁・月・日。例: 261016）_ビュー名.拡張子
    char stamp[16];
    std::time_t now = std::time(nullptr);
    std::strftime(stamp, sizeof(stamp), "%y%m%d_", std::localtime(&now));
    const char* ext = kind == ExportKind::Image ? (png ? ".png" : ".jpg") : kind == ExportKind::Panorama ? ".html" : ".exe";
    return stamp + base + ext;
}

std::string chooseSavePath(void* ownerWindow, ExportKind kind, const std::string& defaultName) {
#ifdef _WIN32
    std::wstring filter, ext;
    switch (kind) {
        case ExportKind::Image: {
            // 既定の名前の拡張子のほうを先に出す（選んだ種類の拡張子が付く）
            std::wstring jpg = widen(tr("FilterJpeg")) + std::wstring(L"\0*.jpg;*.jpeg\0", 14);
            std::wstring png = widen(tr("FilterPng")) + std::wstring(L"\0*.png\0", 7);
            bool pngFirst = defaultName.size() >= 4 && defaultName.compare(defaultName.size() - 4, 4, ".png") == 0;
            filter = pngFirst ? png + jpg : jpg + png;
            ext = pngFirst ? L"png" : L"jpg";
            break;
        }
        case ExportKind::Panorama: filter = widen(tr("FilterHtml")) + std::wstring(L"\0*.html\0", 8); ext = L"html"; break;
        case ExportKind::Viewer: filter = widen(tr("FilterExe")) + std::wstring(L"\0*.exe\0", 7); ext = L"exe"; break;
    }
    filter += L'\0';
    std::wstring buf = widen(defaultName);
    buf.resize(32768, L'\0');
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = static_cast<HWND>(ownerWindow);
    ofn.lpstrFilter = filter.c_str();
    ofn.lpstrFile = &buf[0];
    ofn.nMaxFile = static_cast<DWORD>(buf.size());
    ofn.lpstrDefExt = ext.c_str();
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
    if (!GetSaveFileNameW(&ofn)) return "";
    return narrow(std::wstring(buf.c_str()));
#else
    (void)ownerWindow;
    (void)kind;
    const char* dir = std::getenv("T28_EXPORT_DIR");
    return std::string(dir && *dir ? dir : ".") + "/" + defaultName;
#endif
}

bool chooseColor(void* ownerWindow, uint32_t& color) {
#ifdef _WIN32
    // 「作成した色」の欄は、開くたびに前回の内容を残す
    static COLORREF custom[16] = {RGB(255, 255, 255), RGB(240, 240, 240), RGB(200, 200, 200), RGB(128, 128, 128),
                                  RGB(64, 64, 64),    RGB(0, 0, 0),       RGB(232, 238, 245), RGB(245, 240, 230)};
    CHOOSECOLORW cc{};
    cc.lStructSize = sizeof(cc);
    cc.hwndOwner = static_cast<HWND>(ownerWindow);
    cc.rgbResult = static_cast<COLORREF>(color & 0xFFFFFF);  // COLORREF も 0x00BBGGRR
    cc.lpCustColors = custom;
    cc.Flags = CC_FULLOPEN | CC_RGBINIT;
    if (!ChooseColorW(&cc)) return false;
    color = static_cast<uint32_t>(cc.rgbResult) & 0xFFFFFF;
    return true;
#else
    (void)ownerWindow;
    static const uint32_t presets[] = {0xFFFFFF, 0xF0F0F0, 0x808080, 0x202020, 0xF5EEE8};
    size_t i = 0;
    while (i < 5 && presets[i] != (color & 0xFFFFFF)) ++i;
    color = presets[(i + 1) % 5];
    return true;
#endif
}

bool writePng(const std::string& path, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height) {
    return stbi_write_png(path.c_str(), static_cast<int>(width), static_cast<int>(height), 4, rgba.data(),
                          static_cast<int>(width) * 4) != 0;
}

bool writeImage(const std::string& path, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height) {
    std::string lower = path;
    for (char& ch : lower) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
    if (lower.size() >= 4 && lower.compare(lower.size() - 4, 4, ".png") == 0) return writePng(path, rgba, width, height);
    std::vector<uint8_t> jpeg;
    auto sink = [](void* ctx, void* data, int size) {
        auto* out = static_cast<std::vector<uint8_t>*>(ctx);
        out->insert(out->end(), static_cast<uint8_t*>(data), static_cast<uint8_t*>(data) + size);
    };
    if (!stbi_write_jpg_to_func(sink, &jpeg, static_cast<int>(width), static_cast<int>(height), 4, rgba.data(), 92)) return false;
    return writeAll(path, {&jpeg});
}

bool writePanoramaHtml(const std::string& htmlPath, const std::vector<uint8_t>& rgba, uint32_t width, uint32_t height,
                       const std::string& title) {
    std::vector<uint8_t> jpeg;
    auto sink = [](void* ctx, void* data, int size) {
        auto* out = static_cast<std::vector<uint8_t>*>(ctx);
        out->insert(out->end(), static_cast<uint8_t*>(data), static_cast<uint8_t*>(data) + size);
    };
    if (!stbi_write_jpg_to_func(sink, &jpeg, static_cast<int>(width), static_cast<int>(height), 4, rgba.data(), 92))
        return false;
    // 「画像を保存」で取り出すときのファイル名（HTML と同じ名前の .jpg）
    std::string jpgName = fileNameOf(htmlPath);
    size_t dot = jpgName.find_last_of('.');
    if (dot != std::string::npos) jpgName.resize(dot);
    jpgName += ".jpg";

    const char* lang = currentLanguage() == Lang::US ? "en" : currentLanguage() == Lang::CN ? "zh" : "ja";
    std::string html = kPanoramaHtml;
    html = replaceAll(html, "{{LANG}}", lang);
    html = replaceAll(html, "{{TITLE}}", htmlEscape(title.empty() ? tr("WindowTitle") : title));
    html = replaceAll(html, "{{HINT}}", htmlEscape(tr("PanoramaHint")));
    html = replaceAll(html, "{{NOWEBGL}}", htmlEscape(tr("PanoramaNoWebGL")));
    html = replaceAll(html, "{{SAVE}}", htmlEscape(tr("PanoramaSaveImage")));
    html = replaceAll(html, "{{FILENAME}}", htmlEscape(jpgName));
    html = replaceAll(html, "{{IMAGE}}", base64(jpeg));
    std::vector<uint8_t> bytes(html.begin(), html.end());
    return writeAll(htmlPath, {&bytes});
}

bool writeViewerExe(const std::string& path, const Scene& scene, const Camera& camera, const std::string& title, std::string& error) {
    std::string self = selfExePath();
    std::vector<uint8_t> exe;
    if (self.empty() || !readAll(self, exe)) {
        error = "cannot read " + self;
        return false;
    }
    // この exe 自体がビューアなら、前に付け足したシーンは取り除く
    long old = payloadOffset(exe);
    if (old > 0) exe.resize(static_cast<size_t>(old));

    std::vector<uint8_t> payload(kPayloadMagic, kPayloadMagic + 4);
    uint32_t titleLen = static_cast<uint32_t>(title.size());
    payload.insert(payload.end(), reinterpret_cast<uint8_t*>(&titleLen), reinterpret_cast<uint8_t*>(&titleLen) + 4);
    payload.insert(payload.end(), title.begin(), title.end());
    packScene(scene, camera, payload);

    std::vector<uint8_t> trailer(16);
    uint64_t offset = exe.size();
    memcpy(trailer.data(), &offset, 8);
    memcpy(trailer.data() + 8, kTrailerMagic, 8);
    if (!writeAll(path, {&exe, &payload, &trailer})) {
        error = "cannot write " + path;
        return false;
    }
    return true;
}

bool findEmbeddedScene(std::string& exePath, long& sceneOffset, std::string& title) {
    exePath = selfExePath();
    FILE* f = exePath.empty() ? nullptr : openFile(exePath, "rb");
    if (!f) return false;
    bool found = false;
    char trailer[16];
    if (fseek(f, -16, SEEK_END) == 0 && fread(trailer, 1, 16, f) == 16 && memcmp(trailer + 8, kTrailerMagic, 8) == 0) {
        uint64_t offset = 0;
        memcpy(&offset, trailer, 8);
        char magic[4];
        uint32_t titleLen = 0;
        if (fseek(f, static_cast<long>(offset), SEEK_SET) == 0 && fread(magic, 1, 4, f) == 4 &&
            memcmp(magic, kPayloadMagic, 4) == 0 && fread(&titleLen, 4, 1, f) == 1 && titleLen < 4096) {
            title.resize(titleLen);
            if (titleLen == 0 || fread(&title[0], 1, titleLen, f) == titleLen) {
                sceneOffset = static_cast<long>(offset) + 8 + static_cast<long>(titleLen);
                found = true;
            }
        }
    }
    fclose(f);
    return found;
}

bool loadEmbeddedScene(const std::string& exePath, long sceneOffset, Scene& scene, std::string& error) {
    FILE* f = openFile(exePath, "rb");
    if (!f) {
        error = "cannot open " + exePath;
        return false;
    }
    char magic[4] = {};
    bool packed = fseek(f, sceneOffset, SEEK_SET) == 0 && fread(magic, 1, 4, f) == 4 && memcmp(magic, "T28P", 4) == 0;
    if (!packed) {
        fclose(f);
        return loadScene(exePath, scene, error, sceneOffset);  // 以前の形式（.t28s そのまま）
    }
    // 詰めたシーン: 最後の 16 バイト（付け足した位置の印）の手前までを読んで戻す
    fseek(f, 0, SEEK_END);
    long end = ftell(f) - 16;
    std::vector<uint8_t> data(end > sceneOffset ? static_cast<size_t>(end - sceneOffset) : 0);
    bool ok = !data.empty() && fseek(f, sceneOffset, SEEK_SET) == 0 && fread(data.data(), 1, data.size(), f) == data.size();
    fclose(f);
    if (!ok) {
        error = "cannot read the embedded scene";
        return false;
    }
    return unpackScene(data.data(), data.size(), scene, error);
}

const char* systemLanguage() {
#ifdef _WIN32
    switch (PRIMARYLANGID(GetUserDefaultUILanguage())) {
        case LANG_JAPANESE: return "JP";
        case LANG_CHINESE: return "CN";
        default: return "US";
    }
#else
    const char* l = std::getenv("LANG");
    if (l && strncmp(l, "ja", 2) == 0) return "JP";
    if (l && strncmp(l, "zh", 2) == 0) return "CN";
    return "US";
#endif
}

}  // namespace t28

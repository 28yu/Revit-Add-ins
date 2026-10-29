; ============================================================
;  28 Tools インストーラ（Inno Setup 6 / Unicode 版）
;
;  Revit 2021〜2026 の全バージョン分を 1 本の exe にまとめる。
;  インストール先: C:\ProgramData\Autodesk\Revit\Addins\<ver>\
;    ├── Tools28.addin      ... マニフェスト（Addins 直下）
;    └── 28Tools\*.dll      ... Tools28.dll + ClosedXML 等の依存 DLL
;
;  ビルド（ローカル / CI 共通）:
;    ISCC.exe /DAppVersion=2.6 Installer\28Tools.iss
;    ISCC.exe /DAppVersion=2.6 /DBinRoot=C:\path\to\stage Installer\28Tools.iss
;
;  ⚠ このファイルは UTF-8（BOM 付き）で保存すること。BOM が無いと日本語が化ける。
; ============================================================

; --- ビルド時に上書きできる定義 -------------------------------
; リリースバージョン（CI が /DAppVersion=x.y で渡す）
#ifndef AppVersion
  #define AppVersion "0.0"
#endif
; DLL の置き場所。<BinRoot>\Revit<ver>\*.dll を取り込む（既定はローカルビルド出力）
#ifndef BinRoot
  #define BinRoot "..\bin\Release"
#endif
; .addin の置き場所。<AddinRoot>\<ver>\28Tools\Tools28.addin を取り込む
#ifndef AddinRoot
  #define AddinRoot "..\Packages"
#endif

; 中国語（簡体字）の言語ファイルは Inno Setup 6.5.0 以降でのみ同梱される。
; 古いコンパイラでもビルドが止まらないよう、無ければ日本語＋英語だけで作る。
#if FileExists(AddBackslash(CompilerPath) + "Languages\ChineseSimplified.isl")
  #define HasChinese
#else
  #pragma message "ChineseSimplified.isl が見つからないため中国語を省略します（Inno Setup 6.5 以降で同梱）"
#endif

#define AddinsDir "{commonappdata}\Autodesk\Revit\Addins"

[Setup]
; ⚠ AppId は絶対に変更しないこと。
;   Windows はこの値で「同じアプリの上書き更新」と判定する。変えると旧版と別アプリ扱いになり、
;   「アプリ」一覧に 28 Tools が二重登録され、アンインストールも正しく動かなくなる。
AppId={{02E2B6CC-4A52-4C3B-8EB6-F2A656D9EEB2}
AppName=28 Tools
AppVersion={#AppVersion}
AppVerName=28 Tools v{#AppVersion}
AppPublisher=28 Tools
AppPublisherURL=https://28tools.com/addins.html
AppSupportURL=https://28tools.com/addins.html
AppUpdatesURL=https://28tools.com/addins.html
UninstallDisplayName=28 Tools (Revit Add-in)
; ProgramData に書き込むため管理者権限が必要
PrivilegesRequired=admin
; 実ファイルは Revit の Addins フォルダへ入れる。{app} はアンインストーラの置き場所としてだけ使う
DefaultDirName={commonpf}\28Tools
DisableDirPage=yes
DisableProgramGroupPage=yes
CreateAppDir=yes
Uninstallable=yes
; 前回のタスク選択ではなく「今この PC に入っている Revit」を既定にする
UsePreviousTasks=no
; Revit は CloseApplications では確実に閉じられないため、[Code] で独自に確認する
CloseApplications=no
; ファイルの日時を UTC で記録する。既定(no)だとビルド機(CI=UTC)の時刻を
; 「インストール先 PC の現地時刻」として書き込むため、日本では 9 時間前の日時になる
TimeStampsInUTC=yes
ShowLanguageDialog=auto
OutputDir=Output
OutputBaseFilename=28Tools_Setup_v{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
; 先頭が既定言語（OS の表示言語に一致するものが無い場合に使われる）
Name: "english";  MessagesFile: "compiler:Default.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"
#ifdef HasChinese
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
#endif

[Messages]
; ショートカットを作らないため実際に表示されるのは FinishedLabelNoIcons。念のため両方そろえる
english.FinishedLabel=28 Tools has been installed.%n%nStart (or restart) Revit to load the add-in. A "28 Tools" tab will appear on the ribbon.
english.FinishedLabelNoIcons=28 Tools has been installed.%n%nStart (or restart) Revit to load the add-in. A "28 Tools" tab will appear on the ribbon.
japanese.FinishedLabel=28 Tools のインストールが完了しました。%n%nRevit を起動（再起動）すると読み込まれ、リボンに「28 Tools」タブが表示されます。
japanese.FinishedLabelNoIcons=28 Tools のインストールが完了しました。%n%nRevit を起動（再起動）すると読み込まれ、リボンに「28 Tools」タブが表示されます。
#ifdef HasChinese
chinesesimplified.FinishedLabel=28 Tools 已安装完成。%n%n请启动（或重新启动）Revit 以加载插件，功能区中将显示“28 Tools”选项卡。
chinesesimplified.FinishedLabelNoIcons=28 Tools 已安装完成。%n%n请启动（或重新启动）Revit 以加载插件，功能区中将显示“28 Tools”选项卡。
#endif

[CustomMessages]
english.RevitGroup=Install for the following Revit versions:
english.NoRevitDetected=No installed Revit (2021-2026) was detected.%n%nPlease select the Revit versions to install for manually.
english.NoVersionSelected=Please select at least one Revit version.
english.RevitRunning=Revit is running.%n%nPlease save your work and close Revit, then click [Retry].%n(Add-in files cannot be replaced while Revit is running.)
english.OpenManual=Open the online manual
japanese.RevitGroup=インストールする Revit のバージョン:
japanese.NoRevitDetected=Revit（2021〜2026）が見つかりませんでした。%n%nインストールする Revit のバージョンを手動で選択してください。
japanese.NoVersionSelected=Revit のバージョンを 1 つ以上選択してください。
japanese.RevitRunning=Revit が起動しています。%n%n作業内容を保存して Revit を終了してから［再試行］を押してください。%n（Revit の起動中はアドインのファイルを置き換えられません）
japanese.OpenManual=マニュアル（Web）を開く
#ifdef HasChinese
chinesesimplified.RevitGroup=安装到以下 Revit 版本：
chinesesimplified.NoRevitDetected=未检测到已安装的 Revit（2021-2026）。%n%n请手动选择要安装的 Revit 版本。
chinesesimplified.NoVersionSelected=请至少选择一个 Revit 版本。
chinesesimplified.RevitRunning=Revit 正在运行。%n%n请保存工作并关闭 Revit，然后单击［重试］。%n（Revit 运行期间无法替换插件文件）
chinesesimplified.OpenManual=打开在线手册
#endif

[Tasks]
; 既定のチェック状態は [Code] の CurPageChanged で「インストール済みの Revit だけ」に付け直す
Name: "rvt2021"; Description: "Revit 2021"; GroupDescription: "{cm:RevitGroup}"; Flags: unchecked
Name: "rvt2022"; Description: "Revit 2022"; GroupDescription: "{cm:RevitGroup}"; Flags: unchecked
Name: "rvt2023"; Description: "Revit 2023"; GroupDescription: "{cm:RevitGroup}"; Flags: unchecked
Name: "rvt2024"; Description: "Revit 2024"; GroupDescription: "{cm:RevitGroup}"; Flags: unchecked
Name: "rvt2025"; Description: "Revit 2025"; GroupDescription: "{cm:RevitGroup}"; Flags: unchecked
Name: "rvt2026"; Description: "Revit 2026"; GroupDescription: "{cm:RevitGroup}"; Flags: unchecked

[InstallDelete]
; 旧版（install.bat 時代の初期版）は Tools28.dll を Addins 直下に置いていた。残っていると二重ロードになる
Type: files; Name: "{#AddinsDir}\2021\Tools28.dll"; Tasks: rvt2021
Type: files; Name: "{#AddinsDir}\2022\Tools28.dll"; Tasks: rvt2022
Type: files; Name: "{#AddinsDir}\2023\Tools28.dll"; Tasks: rvt2023
Type: files; Name: "{#AddinsDir}\2024\Tools28.dll"; Tasks: rvt2024
Type: files; Name: "{#AddinsDir}\2025\Tools28.dll"; Tasks: rvt2025
Type: files; Name: "{#AddinsDir}\2026\Tools28.dll"; Tasks: rvt2026
; 依存ライブラリの入れ替わりで古い DLL が残らないよう、28Tools フォルダの DLL を一度消す
Type: files; Name: "{#AddinsDir}\2021\28Tools\*.dll"; Tasks: rvt2021
Type: files; Name: "{#AddinsDir}\2022\28Tools\*.dll"; Tasks: rvt2022
Type: files; Name: "{#AddinsDir}\2023\28Tools\*.dll"; Tasks: rvt2023
Type: files; Name: "{#AddinsDir}\2024\28Tools\*.dll"; Tasks: rvt2024
Type: files; Name: "{#AddinsDir}\2025\28Tools\*.dll"; Tasks: rvt2025
Type: files; Name: "{#AddinsDir}\2026\28Tools\*.dll"; Tasks: rvt2026

[Files]
; 2021〜2024 = .NET Framework 4.8 / 2025〜2026 = .NET 8（ビルド出力をそのまま取り込む）
Source: "{#BinRoot}\Revit2021\*.dll"; DestDir: "{#AddinsDir}\2021\28Tools"; Tasks: rvt2021; Flags: ignoreversion
Source: "{#BinRoot}\Revit2022\*.dll"; DestDir: "{#AddinsDir}\2022\28Tools"; Tasks: rvt2022; Flags: ignoreversion
Source: "{#BinRoot}\Revit2023\*.dll"; DestDir: "{#AddinsDir}\2023\28Tools"; Tasks: rvt2023; Flags: ignoreversion
Source: "{#BinRoot}\Revit2024\*.dll"; DestDir: "{#AddinsDir}\2024\28Tools"; Tasks: rvt2024; Flags: ignoreversion
Source: "{#BinRoot}\Revit2025\*.dll"; DestDir: "{#AddinsDir}\2025\28Tools"; Tasks: rvt2025; Flags: ignoreversion
Source: "{#BinRoot}\Revit2026\*.dll"; DestDir: "{#AddinsDir}\2026\28Tools"; Tasks: rvt2026; Flags: ignoreversion
; リアルタイムレンダリングの描画プログラム（Renderer/ の C++。全バージョン共通の同じ exe）
Source: "{#BinRoot}\Revit2021\Tools28Render.exe"; DestDir: "{#AddinsDir}\2021\28Tools"; Tasks: rvt2021; Flags: ignoreversion
Source: "{#BinRoot}\Revit2022\Tools28Render.exe"; DestDir: "{#AddinsDir}\2022\28Tools"; Tasks: rvt2022; Flags: ignoreversion
Source: "{#BinRoot}\Revit2023\Tools28Render.exe"; DestDir: "{#AddinsDir}\2023\28Tools"; Tasks: rvt2023; Flags: ignoreversion
Source: "{#BinRoot}\Revit2024\Tools28Render.exe"; DestDir: "{#AddinsDir}\2024\28Tools"; Tasks: rvt2024; Flags: ignoreversion
Source: "{#BinRoot}\Revit2025\Tools28Render.exe"; DestDir: "{#AddinsDir}\2025\28Tools"; Tasks: rvt2025; Flags: ignoreversion
Source: "{#BinRoot}\Revit2026\Tools28Render.exe"; DestDir: "{#AddinsDir}\2026\28Tools"; Tasks: rvt2026; Flags: ignoreversion
; 描画プログラムが仕上げに使う Intel Open Image Denoise（CPU 版。約 50MB。全バージョン共通なので、元は 1 か所から取り込み、インストーラには 1 つだけ入る）
Source: "..\Renderer\prebuilt\win-x64\oidn\*"; DestDir: "{#AddinsDir}\2021\28Tools\oidn"; Tasks: rvt2021; Flags: ignoreversion
Source: "..\Renderer\prebuilt\win-x64\oidn\*"; DestDir: "{#AddinsDir}\2022\28Tools\oidn"; Tasks: rvt2022; Flags: ignoreversion
Source: "..\Renderer\prebuilt\win-x64\oidn\*"; DestDir: "{#AddinsDir}\2023\28Tools\oidn"; Tasks: rvt2023; Flags: ignoreversion
Source: "..\Renderer\prebuilt\win-x64\oidn\*"; DestDir: "{#AddinsDir}\2024\28Tools\oidn"; Tasks: rvt2024; Flags: ignoreversion
Source: "..\Renderer\prebuilt\win-x64\oidn\*"; DestDir: "{#AddinsDir}\2025\28Tools\oidn"; Tasks: rvt2025; Flags: ignoreversion
Source: "..\Renderer\prebuilt\win-x64\oidn\*"; DestDir: "{#AddinsDir}\2026\28Tools\oidn"; Tasks: rvt2026; Flags: ignoreversion
; マニフェストは Addins 直下（<Assembly>28Tools\Tools28.dll</Assembly> で 28Tools フォルダを指す）
Source: "{#AddinRoot}\2021\28Tools\Tools28.addin"; DestDir: "{#AddinsDir}\2021"; Tasks: rvt2021; Flags: ignoreversion
Source: "{#AddinRoot}\2022\28Tools\Tools28.addin"; DestDir: "{#AddinsDir}\2022"; Tasks: rvt2022; Flags: ignoreversion
Source: "{#AddinRoot}\2023\28Tools\Tools28.addin"; DestDir: "{#AddinsDir}\2023"; Tasks: rvt2023; Flags: ignoreversion
Source: "{#AddinRoot}\2024\28Tools\Tools28.addin"; DestDir: "{#AddinsDir}\2024"; Tasks: rvt2024; Flags: ignoreversion
Source: "{#AddinRoot}\2025\28Tools\Tools28.addin"; DestDir: "{#AddinsDir}\2025"; Tasks: rvt2025; Flags: ignoreversion
Source: "{#AddinRoot}\2026\28Tools\Tools28.addin"; DestDir: "{#AddinsDir}\2026"; Tasks: rvt2026; Flags: ignoreversion

[UninstallDelete]
; install.bat で入れた分・旧版の直置き DLL も含めて、28 Tools の痕跡を全バージョン分消す
Type: filesandordirs; Name: "{#AddinsDir}\2021\28Tools"
Type: filesandordirs; Name: "{#AddinsDir}\2022\28Tools"
Type: filesandordirs; Name: "{#AddinsDir}\2023\28Tools"
Type: filesandordirs; Name: "{#AddinsDir}\2024\28Tools"
Type: filesandordirs; Name: "{#AddinsDir}\2025\28Tools"
Type: filesandordirs; Name: "{#AddinsDir}\2026\28Tools"
Type: files; Name: "{#AddinsDir}\2021\Tools28.addin"
Type: files; Name: "{#AddinsDir}\2022\Tools28.addin"
Type: files; Name: "{#AddinsDir}\2023\Tools28.addin"
Type: files; Name: "{#AddinsDir}\2024\Tools28.addin"
Type: files; Name: "{#AddinsDir}\2025\Tools28.addin"
Type: files; Name: "{#AddinsDir}\2026\Tools28.addin"
Type: files; Name: "{#AddinsDir}\2021\Tools28.dll"
Type: files; Name: "{#AddinsDir}\2022\Tools28.dll"
Type: files; Name: "{#AddinsDir}\2023\Tools28.dll"
Type: files; Name: "{#AddinsDir}\2024\Tools28.dll"
Type: files; Name: "{#AddinsDir}\2025\Tools28.dll"
Type: files; Name: "{#AddinsDir}\2026\Tools28.dll"

[Run]
; 完了ページの「マニュアルを開く」チェックボックス（既定は未チェック）
Filename: "https://28tools.com/addins.html"; Description: "{cm:OpenManual}"; Flags: postinstall shellexec nowait skipifsilent unchecked

[Code]
const
  FirstRevit = 2021;
  LastRevit = 2026;

var
  TasksInitialized: Boolean;

{ その Revit バージョンがこの PC に入っているか。
  次のどれか 1 つでも当てはまれば「インストール済み」とみなす。
    1. C:\Program Files\Autodesk\Revit <ver>\Revit.exe がある
    2. レジストリ HKLM\SOFTWARE\Autodesk\Revit\<ver> または \Autodesk Revit <ver> がある
    3. C:\ProgramData\Autodesk\Revit\Addins\<ver> フォルダがある（他社アドイン等が作る）
  3 はアンインストール後も残ることがあるが、既定チェックが付くだけなので実害は小さい。 }
function IsRevitInstalled(Ver: Integer): Boolean;
var
  V: String;
begin
  V := IntToStr(Ver);
  Result := DirExists(ExpandConstant('{commonappdata}\Autodesk\Revit\Addins\') + V);
  if (not Result) and IsWin64 then
  begin
    Result := FileExists(ExpandConstant('{commonpf64}\Autodesk\Revit ') + V + '\Revit.exe')
      or RegKeyExists(HKLM64, 'SOFTWARE\Autodesk\Revit\' + V)
      or RegKeyExists(HKLM64, 'SOFTWARE\Autodesk\Revit\Autodesk Revit ' + V);
  end;
end;

{ Revit.exe が起動中か。WMI で調べ、使えない環境では tasklist にフォールバックする }
function IsRevitRunning(): Boolean;
var
  Locator, Service, Procs: Variant;
  ResultCode: Integer;
begin
  Result := False;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Service := Locator.ConnectServer('.', 'root\CIMV2');
    Procs := Service.ExecQuery('SELECT ProcessId FROM Win32_Process WHERE Name = ''Revit.exe''');
    Result := (Procs.Count > 0);
  except
    if Exec(ExpandConstant('{cmd}'),
            '/C tasklist /FI "IMAGENAME eq Revit.exe" /NH | find /I "Revit.exe" >nul',
            '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      Result := (ResultCode = 0);
  end;
end;

{ Revit が終了するまで「再試行 / キャンセル」を出し続ける。キャンセルなら False }
function WaitForRevitClosed(): Boolean;
begin
  Result := True;
  while IsRevitRunning() do
  begin
    if SuppressibleMsgBox(CustomMessage('RevitRunning'), mbError, MB_RETRYCANCEL, IDCANCEL) <> IDRETRY then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

function AnyVersionSelected(): Boolean;
var
  Ver: Integer;
begin
  Result := False;
  for Ver := FirstRevit to LastRevit do
    if WizardIsTaskSelected('rvt' + IntToStr(Ver)) then
    begin
      Result := True;
      Exit;
    end;
end;

function InitializeSetup(): Boolean;
begin
  Result := WaitForRevitClosed();
end;

{ アンインストール時も、Revit が DLL を掴んだままだと削除できないので先に終了してもらう }
function InitializeUninstall(): Boolean;
begin
  Result := WaitForRevitClosed();
end;

procedure CurPageChanged(CurPageID: Integer);
var
  Ver: Integer;
  Found: Boolean;
begin
  { 初回表示時だけ、インストール済みの Revit にチェックを付ける（戻る→進むで上書きしない） }
  if (CurPageID = wpSelectTasks) and (not TasksInitialized) then
  begin
    TasksInitialized := True;
    Found := False;
    for Ver := FirstRevit to LastRevit do
      if IsRevitInstalled(Ver) then
      begin
        WizardSelectTasks('rvt' + IntToStr(Ver));
        Found := True;
      end;
    if not Found then
      SuppressibleMsgBox(CustomMessage('NoRevitDetected'), mbInformation, MB_OK, IDOK);
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = wpSelectTasks) and (not AnyVersionSelected()) then
  begin
    MsgBox(CustomMessage('NoVersionSelected'), mbError, MB_OK);
    Result := False;
  end;
end;

{ 実際のコピー直前にもう一度確認する（ウィザード操作中に Revit を起動された場合に備える）。
  空文字以外を返すとインストールを中止する }
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if not AnyVersionSelected() then
    Result := CustomMessage('NoVersionSelected')
  else if not WaitForRevitClosed() then
    Result := CustomMessage('RevitRunning');
end;

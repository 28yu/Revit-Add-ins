' Tools28 - AutoBuild.ps1 を管理者権限・ウィンドウ非表示で起動するラッパー
' タスクスケジューラやスタートアップから呼び出し用

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
strCommand = "powershell.exe"
strArgs = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strScriptDir & "\AutoBuild.ps1"""

' 管理者権限で起動 (runas) - C:\ProgramData への書き込みに必要
Set objShellApp = CreateObject("Shell.Application")
objShellApp.ShellExecute strCommand, strArgs, strScriptDir, "runas", 0

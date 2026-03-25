' Tools28 - AutoBuild.ps1 を管理者権限でバックグラウンド起動するラッパー
' ダブルクリックで UAC 確認後、非表示で常駐する

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
strCommand = "powershell.exe"
strArgs = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strScriptDir & "\AutoBuild.ps1"""

Set objShellApp = CreateObject("Shell.Application")
objShellApp.ShellExecute strCommand, strArgs, strScriptDir, "runas", 0

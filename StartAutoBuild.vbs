' Tools28 - AutoBuild.ps1 をウィンドウ非表示で起動するラッパー
' タスクスケジューラやスタートアップから呼び出し用
' ※管理者昇格は AutoBuild.ps1 内で自動実行される

Set objShell = CreateObject("WScript.Shell")
strScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strScriptDir & "\AutoBuild.ps1"""
objShell.CurrentDirectory = strScriptDir
objShell.Run strCommand, 0, False

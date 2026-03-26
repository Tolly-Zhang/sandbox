Set sh = CreateObject("WScript.Shell")
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = "powershell -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & "\run.ps1"" -PdfIndex INDEX_VALUE"
sh.Run command, 1

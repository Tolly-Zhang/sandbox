Set sh = CreateObject("WScript.Shell")
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = "powershell -noprofile -executionpolicy bypass -file """ & scriptPath & "\run.ps1"" -FromVbs"
sh.Run command, 1
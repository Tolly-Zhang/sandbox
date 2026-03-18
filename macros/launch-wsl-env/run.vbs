Set objShell = CreateObject("WScript.Shell")
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
command = "powershell -noprofile -windowstyle hidden -executionpolicy bypass -file """ & scriptPath & "\run.ps1"""
objShell.Run command, 0
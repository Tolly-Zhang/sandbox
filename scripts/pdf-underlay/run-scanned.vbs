Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
runPs1 = scriptFolder & "\run.ps1"
runPs1Escaped = Replace(runPs1, "'", "''")

' Create a temporary PowerShell script to avoid complex quoting issues
tempFolder = sh.ExpandEnvironmentStrings("%TEMP%")
tempPs = tempFolder & "\pdf-underlay-launch-" & Replace(CStr(Timer), ".", "-") & ".ps1"

Set tf = fso.CreateTextFile(tempPs, True)
tf.WriteLine "$scriptPath = '" & runPs1Escaped & "'"
tf.WriteLine "$Host.UI.RawUI.WindowTitle = 'PDF Underlay - Scanned'"
tf.WriteLine "Write-Host 'Scanned mode. Paste or drag one or more PDFs into this window, then press Enter.' -ForegroundColor Cyan"
tf.WriteLine "$input = Read-Host 'Paste or drag PDF paths (space or newline separated) and press Enter'"
tf.WriteLine "if ([string]::IsNullOrWhiteSpace($input)) { Write-Host 'No PDFs provided.' -ForegroundColor Yellow }"
tf.WriteLine "else { & $scriptPath -scanned $input }"
tf.WriteLine ""
tf.Close

command = "powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File """ & tempPs & """"
sh.Run command, 1
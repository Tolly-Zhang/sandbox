param(
	[switch]$HiddenChild,
	[switch]$FromVbs
)

$ShowWindow = $true

function Exit-WithError {
	param(
		[string]$Message,
		[int]$Code = 1,
		[System.Exception]$Exception
	)

	Write-Host "ERROR: $Message" -ForegroundColor Red
	if ($Exception) {
		Write-Host $Exception.Message -ForegroundColor DarkRed
	}

	if ($ShowWindow -and $FromVbs) {
		Write-Host ""
		Read-Host "Press Enter to close"
	}

	exit $Code
}

$ConfigPath = Join-Path $PSScriptRoot "config.json"

if (-not (Test-Path -Path $ConfigPath)) {
	Exit-WithError -Message "config.json was not found at $ConfigPath"
}

try {
	$Config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
}
catch {
	Exit-WithError -Message "config.json could not be parsed as JSON" -Exception $_.Exception
}

$Command = [string]$Config.WSLCommand
if ([string]::IsNullOrWhiteSpace($Command)) {
	Exit-WithError -Message "WSLCommand is missing or empty in config.json"
}

$ShowWindow = $false
if ($null -ne $Config.ShowWindow) {
	try {
		$ShowWindow = [System.Convert]::ToBoolean($Config.ShowWindow)
	}
	catch {
		Exit-WithError -Message "ShowWindow must be true or false in config.json" -Exception $_.Exception
	}
}

# Relaunch hidden once when ShowWindow is false.
if (-not $ShowWindow -and -not $HiddenChild) {
	try {
		Start-Process -FilePath "powershell.exe" -WindowStyle Hidden -ArgumentList @(
			"-NoProfile",
			"-ExecutionPolicy",
			"Bypass",
			"-File",
			"`"$PSCommandPath`"",
			"-HiddenChild"
		)
	}
	catch {
		Exit-WithError -Message "Failed to relaunch script in hidden mode" -Exception $_.Exception
	}

	exit 0
}

if (-not (Get-Command wsl.exe -ErrorAction SilentlyContinue)) {
	Exit-WithError -Message "wsl.exe was not found. Install/enable Windows Subsystem for Linux first."
}

try {
	& wsl.exe -e bash -ic $Command
	$WslExitCode = $LASTEXITCODE
}
catch {
	Exit-WithError -Message "Failed to start WSL command." -Exception $_.Exception
}

if ($WslExitCode -ne 0) {
	Exit-WithError -Message "WSL command failed with exit code $WslExitCode. Check WSL setup and WSLCommand in config.json." -Code $WslExitCode
}
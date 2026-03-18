$Config = Get-Content -Path "$PSScriptRoot\config.json" | ConvertFrom-Json
$Command = [string]$Config.WSLCommand

if ([string]::IsNullOrWhiteSpace($Command)) {
	Write-Error "WSLCommand is missing or empty in config.json"
	exit 1
}

wsl.exe -e bash -ic $Command
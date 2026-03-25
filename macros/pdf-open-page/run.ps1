Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$configPath = Join-Path $scriptDir 'config.json'
$config = $null
if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to parse config.json: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit 1
    }
}

$ofd = New-Object System.Windows.Forms.OpenFileDialog
$ofd.Filter = 'PDF files (*.pdf)|*.pdf'
$ofd.Title = 'Select PDF to open'
$ofd.Multiselect = $false
if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit }

$pdf = $ofd.FileName

$pageInput = [Microsoft.VisualBasic.Interaction]::InputBox('Enter page number to open:', 'Open PDF at page', '1')
if ([string]::IsNullOrWhiteSpace($pageInput)) { exit }

[int]$pageNum = 0
if (-not [int]::TryParse($pageInput, [ref]$pageNum) -or $pageNum -lt 1) {
    [System.Windows.Forms.MessageBox]::Show('Please enter a valid page number (>=1).', 'Invalid input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

$viewerPath = $null
$argsTemplate = $null

if ($config -and $config.preferredViewer) {
    $preferred = $config.preferredViewer
    if ($config.viewers -and $config.viewers.$preferred -and $config.viewers.$preferred.path) {
        $p = $config.viewers.$preferred.path
        if (Test-Path $p) {
            $viewerPath = $p
            $argsTemplate = $config.viewers.$preferred.args
        }
    }
}

if (-not $viewerPath -and $config -and $config.viewerPath) {
    if (Test-Path $config.viewerPath) {
        $viewerPath = $config.viewerPath
        $argsTemplate = if ($config.viewerArgsTemplate) { $config.viewerArgsTemplate } else { '-page {page} {file}' }
    }
}

if (-not $viewerPath -and $config -and $config.viewers) {
    foreach ($kv in $config.viewers.PSObject.Properties) {
        $v = $kv.Value
        if ($v.path -and (Test-Path $v.path)) {
            $viewerPath = $v.path
            $argsTemplate = $v.args
            break
        }
    }
}

if (-not $viewerPath) {
    $candidates = @(
        @{ Path = "${env:ProgramFiles}\SumatraPDF\SumatraPDF.exe"; Args = '-page {page} {file}' },
        @{ Path = "${env:ProgramFiles(x86)}\SumatraPDF\SumatraPDF.exe"; Args = '-page {page} {file}' },
        @{ Path = "${env:ProgramFiles(x86)}\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"; Args = '/A "page={page}" {file}' },
        @{ Path = "${env:ProgramFiles}\Adobe\Acrobat DC\Acrobat\Acrobat.exe"; Args = '/A "page={page}" {file}' },
        @{ Path = "${env:ProgramFiles}\Tracker Software\PDF Editor\PDFXEdit.exe"; Args = '/A "page={page}" {file}' },
        @{ Path = "${env:ProgramFiles(x86)}\Tracker Software\PDF Editor\PDFXEdit.exe"; Args = '/A "page={page}" {file}' }
    )
    foreach ($cand in $candidates) {
        if (Test-Path $cand.Path) {
            $viewerPath = $cand.Path
            $argsTemplate = $cand.Args
            break
        }
    }
}

if (-not $viewerPath) {
    [System.Windows.Forms.MessageBox]::Show('No known PDF viewer executable found; opening with default app (may not jump to page).', 'Notice', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Start-Process -FilePath $pdf
    exit
}

$quotedFile = '"' + $pdf + '"'
$argString = $argsTemplate.Replace('{page}', $pageNum.ToString()).Replace('{file}', $quotedFile)

try {
    Start-Process -FilePath $viewerPath -ArgumentList $argString
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to start viewer: $($_.Exception.Message)", 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}
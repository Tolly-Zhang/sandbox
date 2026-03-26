param(
    [ValidateRange(0, 2147483647)]
    [int]$PdfIndex = 0
)

Add-Type -AssemblyName System.Windows.Forms

$script:KeepWindowOpen = $false

function Get-BoolSetting {
    param(
        [object]$Value,
        [bool]$Default = $false
    )

    if ($null -eq $Value) {
        return $Default
    }

    if ($Value -is [bool]) {
        return [bool]$Value
    }

    $parsed = $false
    if ([bool]::TryParse([string]$Value, [ref]$parsed)) {
        return $parsed
    }

    return $Default
}

function Complete-Run {
    if ($script:KeepWindowOpen) {
        Write-Host ''
        Read-Host 'Press Enter to close' | Out-Null
    }

    exit 0
}

function Fail-Run {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [string]$Title = 'Error'
    )

    Write-Host "ERROR: $Message" -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    Write-Host ''
    Read-Host 'Press Enter to close' | Out-Null
    exit 1
}

function ConvertTo-RomanNumeral {
    param(
        [int]$Number,
        [switch]$Lowercase
    )

    if ($Number -le 0) {
        return $Number.ToString()
    }

    $map = @(
        @{ Value = 1000; Symbol = 'M' },
        @{ Value = 900; Symbol = 'CM' },
        @{ Value = 500; Symbol = 'D' },
        @{ Value = 400; Symbol = 'CD' },
        @{ Value = 100; Symbol = 'C' },
        @{ Value = 90; Symbol = 'XC' },
        @{ Value = 50; Symbol = 'L' },
        @{ Value = 40; Symbol = 'XL' },
        @{ Value = 10; Symbol = 'X' },
        @{ Value = 9; Symbol = 'IX' },
        @{ Value = 5; Symbol = 'V' },
        @{ Value = 4; Symbol = 'IV' },
        @{ Value = 1; Symbol = 'I' }
    )

    $remaining = $Number
    $builder = New-Object System.Text.StringBuilder
    foreach ($entry in $map) {
        while ($remaining -ge $entry.Value) {
            [void]$builder.Append($entry.Symbol)
            $remaining -= $entry.Value
        }
    }

    $result = $builder.ToString()
    if ($Lowercase) {
        return $result.ToLowerInvariant()
    }

    return $result
}

function ConvertTo-LetterNumeral {
    param(
        [int]$Number,
        [switch]$Lowercase
    )

    if ($Number -le 0) {
        return $Number.ToString()
    }

    $n = $Number
    $letters = ''
    while ($n -gt 0) {
        $n--
        $letters = [char](65 + ($n % 26)) + $letters
        $n = [math]::Floor($n / 26)
    }

    if ($Lowercase) {
        return $letters.ToLowerInvariant()
    }

    return $letters
}

function Format-PageLabelNumber {
    param(
        [string]$Style,
        [int]$Number
    )

    switch ($Style) {
        'DecimalArabicNumerals' { return $Number.ToString() }
        'UppercaseRomanNumerals' { return ConvertTo-RomanNumeral -Number $Number }
        'LowercaseRomanNumerals' { return ConvertTo-RomanNumeral -Number $Number -Lowercase }
        'UppercaseLetters' { return ConvertTo-LetterNumeral -Number $Number }
        'LowercaseLetters' { return ConvertTo-LetterNumeral -Number $Number -Lowercase }
        'NoNumber' { return '' }
        default { return $Number.ToString() }
    }
}

function Get-LogicalPageLabelCachePath {
    param(
        [string]$PdfPath
    )

    $cacheRoot = Join-Path $env:TEMP 'pdf-open-page-cache'
    if (-not (Test-Path -LiteralPath $cacheRoot)) {
        New-Item -ItemType Directory -Path $cacheRoot -Force | Out-Null
    }

    $hashInput = $PdfPath.ToLowerInvariant()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($hashInput)
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $hashBytes = $sha1.ComputeHash($bytes)
    } finally {
        $sha1.Dispose()
    }

    $hash = [System.BitConverter]::ToString($hashBytes).Replace('-', '').ToLowerInvariant()
    return Join-Path $cacheRoot ("$hash.json")
}

function Get-LogicalPageLabelCache {
    param(
        [string]$PdfPath
    )

    try {
        $fileInfo = Get-Item -LiteralPath $PdfPath -ErrorAction Stop
        $cachePath = Get-LogicalPageLabelCachePath -PdfPath $PdfPath
    } catch {
        return $null
    }

    if (-not (Test-Path -LiteralPath $cachePath)) {
        return $null
    }

    $cacheData = $null
    try {
        $cacheData = Get-Content -LiteralPath $cachePath -Raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return $null
    }

    if (-not $cacheData) {
        return $null
    }

    if ([long]$cacheData.Length -ne [long]$fileInfo.Length) {
        return $null
    }

    $fileWriteTime = $fileInfo.LastWriteTimeUtc.ToString('o')
    if ([string]$cacheData.LastWriteTimeUtc -ne $fileWriteTime) {
        return $null
    }

    if (-not $cacheData.Labels) {
        return $null
    }

    $labels = @{}
    foreach ($entry in $cacheData.Labels) {
        if ($entry -and $entry.Label -and $entry.Page) {
            $labels[[string]$entry.Label] = [int]$entry.Page
        }
    }

    if ($labels.Count -eq 0) {
        return $null
    }

    return [pscustomobject]@{
        NumberOfPages = [int]$cacheData.NumberOfPages
        Labels = $labels
    }
}

function Set-LogicalPageLabelCache {
    param(
        [string]$PdfPath,
        [int]$NumberOfPages,
        [hashtable]$Labels
    )

    try {
        $fileInfo = Get-Item -LiteralPath $PdfPath -ErrorAction Stop
        $cachePath = Get-LogicalPageLabelCachePath -PdfPath $PdfPath
    } catch {
        return
    }

    $labelEntries = @()
    foreach ($key in $Labels.Keys) {
        $labelEntries += [pscustomobject]@{
            Label = $key
            Page = [int]$Labels[$key]
        }
    }

    $cacheObject = [pscustomobject]@{
        Length = [long]$fileInfo.Length
        LastWriteTimeUtc = $fileInfo.LastWriteTimeUtc.ToString('o')
        NumberOfPages = $NumberOfPages
        Labels = $labelEntries
    }

    try {
        $cacheObject | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $cachePath -Encoding UTF8
    } catch {
    }
}

function Get-PdfTkPath {
    param(
        [object]$Config
    )

    $candidates = @()
    if ($Config -and $Config.pdfTkPath) {
        $candidates += [string]$Config.pdfTkPath
    }
    if ($Config -and $Config.tools -and $Config.tools.pdfTkPath) {
        $candidates += [string]$Config.tools.pdfTkPath
    }
    $candidates += 'pdftk'

    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        if (Test-Path -LiteralPath $candidate) {
            try {
                return (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
            } catch {
                return $candidate
            }
        }

        try {
            $resolved = Get-Command $candidate -CommandType Application -ErrorAction Stop
            if ($resolved -and $resolved.Source) {
                return $resolved.Source
            }
        } catch {
        }
    }

    return $null
}

function Get-PdfLogicalPageLabelLookup {
    param(
        [string]$PdfPath,
        [object]$Config
    )

    $cached = Get-LogicalPageLabelCache -PdfPath $PdfPath
    if ($cached) {
        return $cached
    }

    $pdfTkPath = Get-PdfTkPath -Config $Config
    if (-not $pdfTkPath) {
        return $null
    }

    $dumpLines = @()
    try {
        $dumpLines = & $pdfTkPath $PdfPath dump_data_utf8 2>$null
    } catch {
        return $null
    }

    if ($LASTEXITCODE -ne 0 -or -not $dumpLines) {
        return $null
    }

    $numberOfPages = 0
    $segments = @()
    $current = $null

    foreach ($line in $dumpLines) {
        if ($line -match '^NumberOfPages:\s*(\d+)\s*$') {
            $numberOfPages = [int]$Matches[1]
            continue
        }

        if ($line -match '^PageLabelNewIndex:\s*(\d+)\s*$') {
            if ($null -ne $current) {
                $segments += [pscustomobject]$current
            }

            $current = @{
                NewIndex = [int]$Matches[1]
                Start = 1
                Style = 'DecimalArabicNumerals'
                Prefix = ''
            }
            continue
        }

        if ($null -eq $current) {
            continue
        }

        if ($line -match '^PageLabelStart:\s*(\d+)\s*$') {
            $current.Start = [int]$Matches[1]
            continue
        }

        if ($line -match '^PageLabelNumStyle:\s*(.*)$') {
            $current.Style = $Matches[1].Trim()
            continue
        }

        if ($line -match '^PageLabelPrefix:\s*(.*)$') {
            $current.Prefix = $Matches[1]
            continue
        }
    }

    if ($null -ne $current) {
        $segments += [pscustomobject]$current
    }

    if ($segments.Count -eq 0 -or $numberOfPages -lt 1) {
        return $null
    }

    $segments = $segments | Sort-Object -Property NewIndex
    $labels = @{}
    $segmentIndex = 0

    for ($physicalPage = 1; $physicalPage -le $numberOfPages; $physicalPage++) {
        while (($segmentIndex + 1) -lt $segments.Count -and $segments[$segmentIndex + 1].NewIndex -le $physicalPage) {
            $segmentIndex++
        }

        $segment = $segments[$segmentIndex]
        $labelNumber = $segment.Start + ($physicalPage - $segment.NewIndex)
        $numberPart = Format-PageLabelNumber -Style $segment.Style -Number $labelNumber
        $label = "$($segment.Prefix)$numberPart"

        if (-not [string]::IsNullOrWhiteSpace($label) -and -not $labels.ContainsKey($label)) {
            $labels[$label] = $physicalPage
        }
    }

    if ($labels.Count -eq 0) {
        return $null
    }

    Set-LogicalPageLabelCache -PdfPath $PdfPath -NumberOfPages $numberOfPages -Labels $labels

    return [pscustomobject]@{
        NumberOfPages = $numberOfPages
        Labels = $labels
    }
}

function Read-Config {
    param(
        [string]$ConfigPath
    )

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Fail-Run -Message "Failed to parse config.json: $($_.Exception.Message)"
    }
}

function Resolve-ExistingFilePath {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    $expanded = [System.Environment]::ExpandEnvironmentVariables($Path)
    if (-not (Test-Path -LiteralPath $expanded)) {
        return $null
    }

    try {
        return (Resolve-Path -LiteralPath $expanded -ErrorAction Stop).Path
    } catch {
        return $expanded
    }
}

function Get-ConfiguredPdfFiles {
    param(
        [object]$Config
    )

    $configuredPdfFiles = @()

    if (-not ($Config -and $Config.pdfFiles)) {
        return $configuredPdfFiles
    }

    foreach ($rawPath in $Config.pdfFiles) {
        if ($null -eq $rawPath) {
            continue
        }

        $trimmedPath = ([string]$rawPath).Trim()
        if (-not [string]::IsNullOrWhiteSpace($trimmedPath)) {
            $configuredPdfFiles += $trimmedPath
        }
    }

    return $configuredPdfFiles
}

function Select-PdfFromDialog {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'PDF files (*.pdf)|*.pdf'
    $ofd.Title = 'Select PDF to open'
    $ofd.Multiselect = $false

    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return $ofd.FileName
}

function Resolve-TargetPdfPath {
    param(
        [object]$Config,
        [int]$PdfIndex
    )

    if ($PdfIndex -le 0) {
        return Select-PdfFromDialog
    }

    $configuredPdfFiles = Get-ConfiguredPdfFiles -Config $Config
    if ($configuredPdfFiles.Count -eq 0) {
        Fail-Run -Message 'No configured PDFs found in config.json (pdfFiles).' -Title 'Configuration error'
    }

    if ($PdfIndex -gt $configuredPdfFiles.Count) {
        Fail-Run -Message "PdfIndex $PdfIndex is out of range. Valid values: 1-$($configuredPdfFiles.Count)." -Title 'Invalid index'
    }

    $selectedPdf = [System.Environment]::ExpandEnvironmentVariables($configuredPdfFiles[$PdfIndex - 1])
    if (-not [System.IO.Path]::IsPathRooted($selectedPdf)) {
        Fail-Run -Message "Configured path at index $PdfIndex must be an absolute path: $selectedPdf" -Title 'Configuration error'
    }

    $resolvedPdf = Resolve-ExistingFilePath -Path $selectedPdf
    if (-not $resolvedPdf) {
        Fail-Run -Message "Configured PDF does not exist at index ${PdfIndex}: $selectedPdf" -Title 'File not found'
    }

    return $resolvedPdf
}

function Resolve-RequestedPageNumber {
    param(
        [string]$PdfPath,
        [object]$Config
    )

    $logicalPageLookup = Get-PdfLogicalPageLabelLookup -PdfPath $PdfPath -Config $Config
    $pagePrompt = if ($logicalPageLookup) { 'Enter page label/number to open' } else { 'Enter page number to open' }
    $pageInput = Read-Host $pagePrompt

    if ([string]::IsNullOrWhiteSpace($pageInput)) {
        return $null
    }

    [int]$pageNum = 0
    $trimmedPageInput = $pageInput.Trim()

    if ($logicalPageLookup -and $logicalPageLookup.Labels.ContainsKey($trimmedPageInput)) {
        $pageNum = [int]$logicalPageLookup.Labels[$trimmedPageInput]
    } elseif (-not [int]::TryParse($trimmedPageInput, [ref]$pageNum) -or $pageNum -lt 1) {
        $message = if ($logicalPageLookup) {
            'Please enter a valid logical page label or a physical page number (>=1).'
        } else {
            'Please enter a valid page number (>=1).'
        }
        Fail-Run -Message $message -Title 'Invalid input'
    }

    if ($logicalPageLookup -and $logicalPageLookup.NumberOfPages -gt 0 -and $pageNum -gt $logicalPageLookup.NumberOfPages) {
        Fail-Run -Message "Page number is out of range. Max page is $($logicalPageLookup.NumberOfPages)." -Title 'Invalid input'
    }

    return $pageNum
}

function Resolve-ViewerSpec {
    param(
        [object]$Config
    )

    if ($Config -and $Config.preferredViewer -and $Config.viewers) {
        $preferred = [string]$Config.preferredViewer
        $preferredViewerProp = $Config.viewers.PSObject.Properties | Where-Object { $_.Name -eq $preferred } | Select-Object -First 1
        if ($preferredViewerProp -and $preferredViewerProp.Value -and $preferredViewerProp.Value.path) {
            $resolvedPreferredPath = Resolve-ExistingFilePath -Path $preferredViewerProp.Value.path
            if ($resolvedPreferredPath) {
                return [pscustomobject]@{
                    Path = $resolvedPreferredPath
                    ArgsTemplate = $preferredViewerProp.Value.args
                }
            }
        }
    }

    if ($Config -and $Config.viewerPath) {
        $resolvedDirectPath = Resolve-ExistingFilePath -Path $Config.viewerPath
        if ($resolvedDirectPath) {
            return [pscustomobject]@{
                Path = $resolvedDirectPath
                ArgsTemplate = if ($Config.viewerArgsTemplate) { $Config.viewerArgsTemplate } else { '-page {page} {file}' }
            }
        }
    }

    if ($Config -and $Config.viewers) {
        foreach ($kv in $Config.viewers.PSObject.Properties) {
            $v = $kv.Value
            if ($v -and $v.path) {
                $resolvedConfiguredViewerPath = Resolve-ExistingFilePath -Path $v.path
                if ($resolvedConfiguredViewerPath) {
                    return [pscustomobject]@{
                        Path = $resolvedConfiguredViewerPath
                        ArgsTemplate = $v.args
                    }
                }
            }
        }
    }

    $fallbackCandidates = @(
        @{ Path = "${env:ProgramFiles}\SumatraPDF\SumatraPDF.exe"; Args = '-page {page} {file}' },
        @{ Path = "${env:ProgramFiles(x86)}\SumatraPDF\SumatraPDF.exe"; Args = '-page {page} {file}' },
        @{ Path = "${env:ProgramFiles(x86)}\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"; Args = '/A "page={page}" {file}' },
        @{ Path = "${env:ProgramFiles}\Adobe\Acrobat DC\Acrobat\Acrobat.exe"; Args = '/A "page={page}" {file}' },
        @{ Path = "${env:ProgramFiles}\Tracker Software\PDF Editor\PDFXEdit.exe"; Args = '/A "page={page}" {file}' },
        @{ Path = "${env:ProgramFiles(x86)}\Tracker Software\PDF Editor\PDFXEdit.exe"; Args = '/A "page={page}" {file}' }
    )

    foreach ($candidate in $fallbackCandidates) {
        $resolvedFallbackPath = Resolve-ExistingFilePath -Path $candidate.Path
        if ($resolvedFallbackPath) {
            return [pscustomobject]@{
                Path = $resolvedFallbackPath
                ArgsTemplate = $candidate.Args
            }
        }
    }

    return $null
}

function Start-PdfAtPage {
    param(
        [string]$ViewerPath,
        [string]$ArgsTemplate,
        [string]$PdfPath,
        [int]$PageNumber
    )

    $quotedFile = '"' + $PdfPath + '"'
    $argString = $ArgsTemplate.Replace('{page}', $PageNumber.ToString()).Replace('{file}', $quotedFile)

    try {
        Start-Process -FilePath $ViewerPath -ArgumentList $argString
    } catch {
        Fail-Run -Message "Failed to start viewer: $($_.Exception.Message)"
    }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$configPath = Join-Path $scriptDir 'config.json'
$config = Read-Config -ConfigPath $configPath

if ($config -and $null -ne $config.keepWindowOpen) {
    $script:KeepWindowOpen = Get-BoolSetting -Value $config.keepWindowOpen -Default $false
}

$pdf = Resolve-TargetPdfPath -Config $config -PdfIndex $PdfIndex
if (-not $pdf) {
    Complete-Run
}

$pageNum = Resolve-RequestedPageNumber -PdfPath $pdf -Config $config
if ($null -eq $pageNum) {
    Complete-Run
}

$viewerSpec = Resolve-ViewerSpec -Config $config
if (-not $viewerSpec) {
    [System.Windows.Forms.MessageBox]::Show('No known PDF viewer executable found; opening with default app (may not jump to page).', 'Notice', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Start-Process -FilePath $pdf
    Complete-Run
}

Start-PdfAtPage -ViewerPath $viewerSpec.Path -ArgsTemplate $viewerSpec.ArgsTemplate -PdfPath $pdf -PageNumber $pageNum
Complete-Run
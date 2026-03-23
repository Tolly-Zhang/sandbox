param(
    [Parameter(Mandatory=$true, Position=0, ValueFromRemainingArguments=$true)]
    [string[]]$PdfPaths,

    [switch]$Scanned,

    [int]$Dpi = -1,
    [int]$Fuzz = -1,

    [switch]$NoBackup,
    [switch]$KeepTemp
)

$ErrorActionPreference = "Stop"

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

    exit $Code
}

function Get-ConfigProperty {
    param(
        [object]$Object,
        [string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) {
        return $Default
    }

    if ($Object.PSObject.Properties.Name -contains $Name) {
        return $Object.$Name
    }

    return $Default
}

function Convert-ToIntValue {
    param(
        $Value,
        [int]$Default,
        [string]$Name
    )

    if ($null -eq $Value) {
        return $Default
    }

    if ($Value -is [int]) {
        return [int]$Value
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $Default
    }

    $parsed = 0
    if ([int]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    throw "$Name must be an integer value."
}

function Convert-ToBooleanValue {
    param(
        $Value,
        [bool]$Default,
        [string]$Name
    )

    if ($null -eq $Value) {
        return $Default
    }

    if ($Value -is [bool]) {
        return [bool]$Value
    }

    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $Default
    }

    switch ($text) {
        "true" { return $true }
        "1" { return $true }
        "yes" { return $true }
        "y" { return $true }
        "on" { return $true }

        "false" { return $false }
        "0" { return $false }
        "no" { return $false }
        "n" { return $false }
        "off" { return $false }

        default { throw "$Name must be a boolean value (true/false)." }
    }
}

# Expand inputs that may contain multiple paths pasted together (newlines, quoted tokens, etc.)
function Expand-PdfPaths {
    param(
        [string[]]$Inputs
    )

    $expanded = @()
    if ($null -eq $Inputs) { return $expanded }

    foreach ($item in $Inputs) {
        if ($null -eq $item) { continue }
        $text = [string]$item
        if ([string]::IsNullOrWhiteSpace($text)) { continue }

        # Normalize line endings and split into lines
        $normalized = $text -replace "`r", ""
        $lines = $normalized -split "`n"

        foreach ($line in $lines) {
            $lineTrim = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($lineTrim)) { continue }

            # Tokenize: match double-quoted, single-quoted, or non-whitespace tokens
            $tokenPattern = '(?:"([^"]+)"|''([^'']+)''|(\S+))'
            $matches = [regex]::Matches($lineTrim, $tokenPattern)

            if ($matches.Count -gt 0) {
                foreach ($m in $matches) {
                    $val = $m.Groups[1].Value
                    if ([string]::IsNullOrEmpty($val)) { $val = $m.Groups[2].Value }
                    if ([string]::IsNullOrEmpty($val)) { $val = $m.Groups[3].Value }
                    if (-not [string]::IsNullOrWhiteSpace($val)) {
                        $clean = $val.Trim().Trim([char]34, [char]39).TrimEnd(',',';')
                        if (-not [string]::IsNullOrWhiteSpace($clean)) {
                            $expanded += $clean
                        }
                    }
                }
            }
            else {
                $clean = $lineTrim.Trim().Trim([char]34, [char]39).TrimEnd(',',';')
                if (-not [string]::IsNullOrWhiteSpace($clean)) {
                    $expanded += $clean
                }
            }
        }
    }

    return $expanded
}

function Resolve-ConfiguredTool {
    param(
        [string]$ConfiguredPath,
        [string]$FallbackCommand,
        [string]$FriendlyName,
        [switch]$Required
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredPath)) {
        $expanded = [Environment]::ExpandEnvironmentVariables($ConfiguredPath)

        if (Test-Path -LiteralPath $expanded) {
            return (Resolve-Path -LiteralPath $expanded).Path
        }

        $commandFromConfigured = Get-Command $expanded -ErrorAction SilentlyContinue
        if ($commandFromConfigured) {
            return $commandFromConfigured.Source
        }

        if ($Required) {
            throw "$FriendlyName was not found at configured path: $ConfiguredPath"
        }

        return $null
    }

    $fallback = Get-Command $FallbackCommand -ErrorAction SilentlyContinue
    if ($fallback) {
        return $fallback.Source
    }

    if ($Required) {
        throw "$FriendlyName was not found. Set its path in config.json"
    }

    return $null
}

function Resolve-UnderlayPath {
    param([string]$UnderlayFileName)

    if ([string]::IsNullOrWhiteSpace($UnderlayFileName)) {
        throw "UnderlayFileName is missing in config.json"
    }

    $expanded = [Environment]::ExpandEnvironmentVariables($UnderlayFileName)
    $candidate = $expanded

    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $candidate = Join-Path $PSScriptRoot $candidate
    }

    if (-not (Test-Path -LiteralPath $candidate)) {
        throw "Underlay reference file was not found: $candidate"
    }

    return (Resolve-Path -LiteralPath $candidate).Path
}

function Invoke-Tool {
    param(
        [string]$Tool,
        [string[]]$Arguments,
        [string]$FailureMessage,
        [switch]$Quiet,
        [switch]$SuppressStdErr
    )

    if ($Quiet) {
        if ($SuppressStdErr) {
            & $Tool @Arguments 2>$null | Out-Null
        }
        else {
            & $Tool @Arguments | Out-Null
        }
    }
    else {
        if ($SuppressStdErr) {
            & $Tool @Arguments 2>$null
        }
        else {
            & $Tool @Arguments
        }
    }

    $exit = $LASTEXITCODE
    if ($exit -ne 0) {
        try {
            $exeName = [System.IO.Path]::GetFileName($Tool).ToLowerInvariant()
        }
        catch {
            $exeName = $Tool.ToLowerInvariant()
        }

        if ($exeName -match '^qpdf(\.exe)?$' -and $exit -eq 3) {
            Write-Host "WARNING: $FailureMessage (qpdf reported warnings - exit code 3)" -ForegroundColor Yellow
            return
        }

        throw "$FailureMessage (exit code $exit)"
    }
}

function Backup-OriginalPdf {
    param(
        [string]$ResolvedTargetPdf,
        [bool]$CreateBackup
    )

    if (-not $CreateBackup) {
        return
    }

    Copy-Item -LiteralPath $ResolvedTargetPdf -Destination ($ResolvedTargetPdf + ".bak") -Force
}

function Resolve-RunSettings {
    param(
        [object]$Config,
        [int]$Dpi,
        [int]$Fuzz,
        [switch]$NoBackup,
        [switch]$KeepTemp
    )

    $defaults = Get-ConfigProperty -Object $Config -Name "Defaults"

    $effectiveDpi = $Dpi
    if ($effectiveDpi -lt 0) {
        $effectiveDpi = Convert-ToIntValue -Value (Get-ConfigProperty -Object $defaults -Name "Dpi" -Default 0) -Default 0 -Name "Defaults.Dpi"
    }

    $effectiveFuzz = $Fuzz
    if ($effectiveFuzz -lt 0) {
        $effectiveFuzz = Convert-ToIntValue -Value (Get-ConfigProperty -Object $defaults -Name "Fuzz" -Default 8) -Default 8 -Name "Defaults.Fuzz"
    }

    $createBackup = Convert-ToBooleanValue -Value (Get-ConfigProperty -Object $defaults -Name "CreateBackup" -Default $true) -Default $true -Name "Defaults.CreateBackup"
    if ($NoBackup) {
        $createBackup = $false
    }

    $keepTempFinal = Convert-ToBooleanValue -Value (Get-ConfigProperty -Object $defaults -Name "KeepTemp" -Default $false) -Default $false -Name "Defaults.KeepTemp"
    if ($KeepTemp) {
        $keepTempFinal = $true
    }

    $underlayPdf = Resolve-UnderlayPath -UnderlayFileName ([string](Get-ConfigProperty -Object $Config -Name "UnderlayFileName"))

    return [pscustomobject]@{
        UnderlayPdf = $underlayPdf
        Dpi = $effectiveDpi
        Fuzz = $effectiveFuzz
        CreateBackup = $createBackup
        KeepTemp = $keepTempFinal
    }
}

function Resolve-ModeTools {
    param(
        [object]$ToolsConfig,
        [switch]$Scanned
    )

    if ($Scanned) {
        return [pscustomobject]@{
            PdfToPpm = Resolve-ConfiguredTool -ConfiguredPath ([string](Get-ConfigProperty -Object $ToolsConfig -Name "PdfToPpmPath")) -FallbackCommand "pdftoppm" -FriendlyName "pdftoppm" -Required
            PdfImages = Resolve-ConfiguredTool -ConfiguredPath ([string](Get-ConfigProperty -Object $ToolsConfig -Name "PdfImagesPath")) -FallbackCommand "pdfimages" -FriendlyName "pdfimages" -Required
            Magick = Resolve-ConfiguredTool -ConfiguredPath ([string](Get-ConfigProperty -Object $ToolsConfig -Name "MagickPath")) -FallbackCommand "magick" -FriendlyName "magick" -Required
            PdfTk = Resolve-ConfiguredTool -ConfiguredPath ([string](Get-ConfigProperty -Object $ToolsConfig -Name "PdfTkPath")) -FallbackCommand "pdftk" -FriendlyName "pdftk" -Required
        }
    }

    return [pscustomobject]@{
        Qpdf = Resolve-ConfiguredTool -ConfiguredPath ([string](Get-ConfigProperty -Object $ToolsConfig -Name "QpdfPath")) -FallbackCommand "qpdf" -FriendlyName "qpdf" -Required
    }
}

function Detect-DpiFromPdf {
    param(
        [string]$PdfImages,
        [string]$PdfPath,
        [int]$Fallback
    )

    if (-not $PdfImages) {
        return $Fallback
    }

    try {
        $imageInfo = & $PdfImages -list -f 1 -l 1 $PdfPath 2>$null | Out-String
    }
    catch {
        return $Fallback
    }

    $xDpi = 0
    $yDpi = 0

    if ($imageInfo -match "x-ppi:\s*(\d+)") {
        $xDpi = [int]$Matches[1]
    }
    if ($imageInfo -match "y-ppi:\s*(\d+)") {
        $yDpi = [int]$Matches[1]
    }

    if ($xDpi -gt 0 -and $yDpi -gt 0) {
        return [Math]::Max($xDpi, $yDpi)
    }

    return $Fallback
}

function Invoke-SimpleUnderlay {
    param(
        [string]$ResolvedTargetPdf,
        [string]$Qpdf,
        [string]$UnderlayPdf
    )

    $targetDir = Split-Path -Path $ResolvedTargetPdf -Parent
    $targetBase = [System.IO.Path]::GetFileNameWithoutExtension($ResolvedTargetPdf)
    $targetExt = [System.IO.Path]::GetExtension($ResolvedTargetPdf)
    $tempOut = Join-Path $targetDir ("{0}-gridtmp{1}" -f $targetBase, $targetExt)

    Invoke-Tool -Tool $Qpdf -Arguments @($ResolvedTargetPdf, "--underlay", $UnderlayPdf, "--repeat=1", "--", $tempOut) -FailureMessage "qpdf failed for '$ResolvedTargetPdf'"

    if (-not (Test-Path -LiteralPath $tempOut)) {
        throw "qpdf did not create output file: $tempOut"
    }

    Move-Item -LiteralPath $tempOut -Destination $ResolvedTargetPdf -Force
}

function Invoke-ScannedUnderlay {
    param(
        [string]$ResolvedTargetPdf,
        [string]$UnderlayPdf,
        [string]$PdfToPpm,
        [string]$PdfImages,
        [string]$Magick,
        [string]$PdfTk,
        [int]$ConfiguredDpi,
        [int]$Fuzz,
        [bool]$KeepTemp
    )

    $effectiveDpi = $ConfiguredDpi
    if ($effectiveDpi -le 0) {
        $effectiveDpi = Detect-DpiFromPdf -PdfImages $PdfImages -PdfPath $ResolvedTargetPdf -Fallback 600
    }

    $work = Join-Path $env:TEMP ("pdfus_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $work | Out-Null

    $pushed = $false

    try {
        Push-Location $work
        $pushed = $true

        Invoke-Tool -Tool $PdfToPpm -Arguments @("-png", "-r", [string]$effectiveDpi, $ResolvedTargetPdf, "overlay") -FailureMessage "pdftoppm failed for '$ResolvedTargetPdf'" -Quiet

        $overlayPages = Get-ChildItem -LiteralPath $work -Filter "overlay-*.png" | Sort-Object Name
        if ($overlayPages.Count -eq 0) {
            throw "No overlay pages extracted from target PDF."
        }

        foreach ($ov in $overlayPages) {
            Invoke-Tool -Tool $Magick -Arguments @($ov.Name, "-fuzz", ("{0}%" -f $Fuzz), "-transparent", "white", ("t_" + $ov.Name)) -FailureMessage "ImageMagick transparency conversion failed for '$($ov.Name)'" -Quiet
        }

        $transOverlayPages = Get-ChildItem -LiteralPath $work -Filter "t_overlay-*.png" | Sort-Object Name
        if ($transOverlayPages.Count -eq 0) {
            throw "No transparent overlay images produced."
        }

        foreach ($png in $transOverlayPages) {
            $pdfName = $png.BaseName + ".pdf"
            Invoke-Tool -Tool $Magick -Arguments @($png.Name, "-background", "none", "-alpha", "Background", $pdfName) -FailureMessage "ImageMagick PDF conversion failed for '$($png.Name)'" -Quiet
        }

        $overlayPdfs = Get-ChildItem -LiteralPath $work -Filter "t_overlay-*.pdf" | Sort-Object Name
        if ($overlayPdfs.Count -eq 0) {
            throw "No overlay PDFs created."
        }

        $underlayInfo = & $PdfTk $UnderlayPdf dump_data 2>$null | Out-String
        if ($LASTEXITCODE -ne 0) {
            throw "pdftk dump_data failed for underlay PDF."
        }

        $underlayPageCount = 1
        if ($underlayInfo -match "NumberOfPages:\s+(\d+)") {
            $underlayPageCount = [int]$Matches[1]
        }

        for ($i = 0; $i -lt $overlayPdfs.Count; $i++) {
            $pageNum = $i + 1
            $overlayPdf = $overlayPdfs[$i].Name

            $underlayPageNum = (($i % $underlayPageCount) + 1)
            $underlayPageFile = "underlay_page_$pageNum.pdf"
            Invoke-Tool -Tool $PdfTk -Arguments @($UnderlayPdf, "cat", [string]$underlayPageNum, "output", $underlayPageFile) -FailureMessage "pdftk page extraction failed for output page $pageNum" -Quiet -SuppressStdErr

            $resultPageFile = "result_page_{0:D4}.pdf" -f $pageNum
            Invoke-Tool -Tool $PdfTk -Arguments @($underlayPageFile, "stamp", $overlayPdf, "output", $resultPageFile) -FailureMessage "pdftk stamp failed for output page $pageNum" -Quiet -SuppressStdErr
        }

        $resultPdfs = Get-ChildItem -LiteralPath $work -Filter "result_page_*.pdf" | Sort-Object Name
        if ($resultPdfs.Count -eq 0) {
            throw "No result pages were generated."
        }

        $resultPdfList = $resultPdfs | ForEach-Object { $_.Name }
        $mergeArgs = @()
        $mergeArgs += $resultPdfList
        $mergeArgs += "cat", "output", "result.pdf"
        Invoke-Tool -Tool $PdfTk -Arguments $mergeArgs -FailureMessage "pdftk final merge failed" -Quiet -SuppressStdErr

        Move-Item -LiteralPath (Join-Path $work "result.pdf") -Destination $ResolvedTargetPdf -Force
    }
    finally {
        if ($pushed) {
            Pop-Location
        }

        if ($KeepTemp) {
            Write-Host "Temp kept at: $work" -ForegroundColor Yellow
        }
        else {
            Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-PdfProcessing {
    param(
        [string]$ResolvedInput,
        [switch]$Scanned,
        [pscustomobject]$Settings,
        [pscustomobject]$ModeTools
    )

    Backup-OriginalPdf -ResolvedTargetPdf $ResolvedInput -CreateBackup $Settings.CreateBackup

    if ($Scanned) {
        Invoke-ScannedUnderlay -ResolvedTargetPdf $ResolvedInput -UnderlayPdf $Settings.UnderlayPdf -PdfToPpm $ModeTools.PdfToPpm -PdfImages $ModeTools.PdfImages -Magick $ModeTools.Magick -PdfTk $ModeTools.PdfTk -ConfiguredDpi $Settings.Dpi -Fuzz $Settings.Fuzz -KeepTemp $Settings.KeepTemp
    }
    else {
        Invoke-SimpleUnderlay -ResolvedTargetPdf $ResolvedInput -Qpdf $ModeTools.Qpdf -UnderlayPdf $Settings.UnderlayPdf
    }
}

$configPath = Join-Path $PSScriptRoot "config.json"
if (-not (Test-Path -LiteralPath $configPath)) {
    Exit-WithError -Message "config.json was not found at $configPath"
}

# If any provided argument contains multiple paths (pasted together), expand them now
try {
    $PdfPaths = Expand-PdfPaths -Inputs $PdfPaths
}
catch {
    Exit-WithError -Message "Failed to parse provided PDF paths" -Exception $_.Exception
}

try {
    $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
}
catch {
    Exit-WithError -Message "config.json could not be parsed as JSON" -Exception $_.Exception
}

try {
    $toolsConfig = Get-ConfigProperty -Object $config -Name "Tools"
    $settings = Resolve-RunSettings -Config $config -Dpi $Dpi -Fuzz $Fuzz -NoBackup:$NoBackup -KeepTemp:$KeepTemp
    $modeTools = Resolve-ModeTools -ToolsConfig $toolsConfig -Scanned:$Scanned

    $succeeded = 0
    $failed = 0

    foreach ($inputPath in $PdfPaths) {
        try {
            if ([string]::IsNullOrWhiteSpace($inputPath)) {
                throw "An empty PDF path was provided."
            }

            if (-not (Test-Path -LiteralPath $inputPath)) {
                throw "File not found: $inputPath"
            }

            $resolvedInput = (Resolve-Path -LiteralPath $inputPath).Path
            Write-Host "Processing: $resolvedInput" -ForegroundColor Cyan

            Invoke-PdfProcessing -ResolvedInput $resolvedInput -Scanned:$Scanned -Settings $settings -ModeTools $modeTools

            $succeeded++
            Write-Host "Done: $resolvedInput" -ForegroundColor Green
        }
        catch {
            $failed++
            Write-Host "Failed: $inputPath" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor DarkRed
        }
    }

    Write-Host ""
    Write-Host "Summary: succeeded=$succeeded failed=$failed" -ForegroundColor Gray

    if ($failed -gt 0) {
        exit 1
    }
}
catch {
    Exit-WithError -Message "Execution failed" -Exception $_.Exception
}

exit 0

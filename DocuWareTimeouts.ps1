Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================
# SETTINGS
# ============================

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

# Desired DAL attributes (KBA-34459)
$desiredAttributes = @{
    AutoTempTableListSize  = '200'
    LongCommandTimeout     = '1800'
    UsePerformanceCounters = 'true'
    UseAutoParameters      = 'true'
    ConnectionTimeout      = '600'
    CommandTimeout         = '600'
}

# Roots to scan
$searchRoots = @(
    'C:\Program Files (x86)\DocuWare\Authentication Server',
    'C:\Program Files\DocuWare\Web',
    'C:\Program Files\DocuWare\Background Process Service',
    'C:\Program Files\DocuWare\Server Manager',
    'C:\Program Files (x86)\DocuWare\Setup Components',
    'C:\Program Files (x86)\DocuWare\Power Tools'
)

# Heuristic: pick first <dataSettings ...> whose following content contains <dataProviders>
$LookAheadChars = 20000

# Encoding step: convert to UTF-8 (NO BOM) only if BOM exists
$ConvertToUtf8NoBom_OnlyIfBomExists = $true

# Logs
$logDirectory = 'C:\Temp\DocuWare-DAL-Config-Logs'
New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
$csvPath = Join-Path $logDirectory "DAL_Config_ZeroReformat_$timestamp.csv"


# ============================
# HELPERS
# ============================

function Get-BomType {
    param([byte[]]$Bytes)

    if ($Bytes.Length -ge 4) {
        if ($Bytes[0] -eq 0x00 -and $Bytes[1] -eq 0x00 -and $Bytes[2] -eq 0xFE -and $Bytes[3] -eq 0xFF) { return 'UTF32BE' }
        if ($Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE -and $Bytes[2] -eq 0x00 -and $Bytes[3] -eq 0x00) { return 'UTF32LE' }
    }
    if ($Bytes.Length -ge 3) {
        if ($Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { return 'UTF8BOM' }
    }
    if ($Bytes.Length -ge 2) {
        if ($Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) { return 'UTF16LE' }
        if ($Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) { return 'UTF16BE' }
    }
    return $null
}

function Read-TextWithEncodingDetection {
    param([Parameter(Mandatory)] [string] $Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $bomType = Get-BomType -Bytes $bytes
    $hasBom = [bool]$bomType

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        # detectEncodingFromByteOrderMarks = $true
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8, $true)
        try {
            $text = $sr.ReadToEnd()
            $enc  = $sr.CurrentEncoding
        } finally { $sr.Dispose() }
    } finally { $fs.Dispose() }

    [PSCustomObject]@{
        Text     = $text
        Encoding = $enc
        HasBom   = $hasBom
        BomType  = $bomType
    }
}

function Write-TextPreserveEncoding {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Text,
        [Parameter(Mandatory)] [System.Text.Encoding] $Encoding,
        [Parameter(Mandatory)] [bool] $HasBom
    )

    $body = $Encoding.GetBytes($Text)

    if ($HasBom -and $Encoding.GetPreamble().Length -gt 0) {
        $bom = $Encoding.GetPreamble()
        $out = New-Object byte[] ($bom.Length + $body.Length)
        [System.Buffer]::BlockCopy($bom, 0, $out, 0, $bom.Length)
        [System.Buffer]::BlockCopy($body, 0, $out, $bom.Length, $body.Length)
        [System.IO.File]::WriteAllBytes($Path, $out)
    } else {
        [System.IO.File]::WriteAllBytes($Path, $body)
    }
}

function Write-Utf8NoBom {
    param([Parameter(Mandatory)] [string] $Path, [Parameter(Mandatory)] [string] $Text)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false) # false => no BOM
    [System.IO.File]::WriteAllBytes($Path, $utf8NoBom.GetBytes($Text))
}

function Mask-XmlComments {
    param([Parameter(Mandatory)] [string] $Text)
    # Replace <!-- ... --> with same-length spaces so indices remain aligned
    $rx = New-Object System.Text.RegularExpressions.Regex('<!--.*?-->', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $rx.Replace($Text, { param($m) (' ' * $m.Value.Length) })
}

function Find-FirstActiveDataSettingsTag {
    param(
        [Parameter(Mandatory)] [string] $Text,
        [int] $LookAheadChars = 20000
    )

    $masked = Mask-XmlComments -Text $Text

    # Match full start tag: <dataSettings ...>
    $rxTag = New-Object System.Text.RegularExpressions.Regex(
        '<\s*dataSettings\b[^>]*>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    $matches = $rxTag.Matches($masked)
    if ($matches.Count -eq 0) { return $null }

    # Prefer the first that looks active (followed by <dataProviders>)
    foreach ($m in $matches) {
        $after = $m.Index + $m.Length
        $len = [Math]::Min($LookAheadChars, ($masked.Length - $after))
        $lookAhead = if ($len -gt 0) { $masked.Substring($after, $len) } else { '' }

        if ($lookAhead -match '(?i)<\s*dataProviders\b') {
            return [PSCustomObject]@{ Start = $m.Index; End = ($m.Index + $m.Length - 1) }
        }
    }

    # Fallback: first uncommented <dataSettings ...>
    $m0 = $matches[0]
    return [PSCustomObject]@{ Start = $m0.Index; End = ($m0.Index + $m0.Length - 1) }
}

function Patch-DataSettingsTag {
    param(
        [Parameter(Mandatory)] [string] $TagText,
        [Parameter(Mandatory)] [hashtable] $DesiredAttributes
    )

    $changed = $false
    $newTag = $TagText

    foreach ($key in $DesiredAttributes.Keys) {
        $desiredValue = [string]$DesiredAttributes[$key]
        $escapedKey = [System.Text.RegularExpressions.Regex]::Escape($key)

        # Match existing attr: key="..." or key='...'
        $pattern = "(?i)(\s$escapedKey\s*=\s*)(['""])(.*?)(\2)"
        $m = [System.Text.RegularExpressions.Regex]::Match($newTag, $pattern)

        if ($m.Success) {
            $curVal = $m.Groups[3].Value
            if ($curVal -ne $desiredValue) { $changed = $true }

            $prefix = $m.Groups[1].Value
            $quote  = $m.Groups[2].Value
            $replacement = $prefix + $quote + $desiredValue + $quote

            $newTag = $newTag.Remove($m.Index, $m.Length).Insert($m.Index, $replacement)
        }
        else {
            # Add missing attr: insert right before closing > or />
            $ins = " $key=`"$desiredValue`""
            $newTag2 = [System.Text.RegularExpressions.Regex]::Replace($newTag, '(\s*/?>)\s*$', ($ins + '$1'), 1)
            if ($newTag2 -ne $newTag) {
                $newTag = $newTag2
                $changed = $true
            }
        }
    }

    [PSCustomObject]@{ NewTag = $newTag; Changed = $changed }
}

function Convert-ToUtf8NoBom-IfBom {
    param([Parameter(Mandatory)] [string] $Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $bomType = Get-BomType -Bytes $bytes
    if (-not $bomType) {
        return [PSCustomObject]@{ Converted = $false; BomType = $null; FromEnc = $null }
    }

    $file = Read-TextWithEncodingDetection -Path $Path
    Write-Utf8NoBom -Path $Path -Text $file.Text
    return [PSCustomObject]@{ Converted = $true; BomType = $bomType; FromEnc = $file.Encoding.WebName }
}


# ============================
# DISCOVERY (DLL adjacency)
# ============================

Write-Host "Discovering DocuWare DAL config files..." -ForegroundColor Cyan

$rootsToScan = $searchRoots | Where-Object { Test-Path $_ } | Sort-Object -Unique
Write-Host ("Roots to scan: {0}" -f $rootsToScan.Count)

$cfgSet = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($root in $rootsToScan) {
    Write-Host "  Scanning $root" -ForegroundColor DarkCyan
    try {
        foreach ($dllPath in [System.IO.Directory]::EnumerateFiles($root, 'DocuWare.DAL.dll', [System.IO.SearchOption]::AllDirectories)) {
            $configPath = "$dllPath.config"
            if ([System.IO.File]::Exists($configPath)) {
                [void]$cfgSet.Add($configPath)
            }
        }
    }
    catch {
        Write-Host "    Warning (skipped some paths): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$configPaths = @($cfgSet) | Sort-Object
Write-Host ("Found {0} DAL config file(s)." -f $configPaths.Count) -ForegroundColor Green

if (-not $configPaths.Count) {
    Write-Host "No matching DocuWare.DAL.dll.config files found. Exiting." -ForegroundColor Yellow
    return
}


# ============================
# PROCESS (zero reformat)
# ============================

$log = New-Object System.Collections.Generic.List[psobject]

foreach ($path in $configPaths) {

    Write-Host "Processing $path"

    $entry = [PSCustomObject]@{
        Path             = $path
        PatchStatus      = $null     # Patched / AlreadyCompliant / MissingDataSettings / Error
        EncodingStatus   = $null     # Converted_ToUtf8NoBOM / Skipped_NoBOM / NotRequested / Error
        BomType          = $null
        FromEncoding     = $null
        BackupCreated    = $false
        Timestamp        = Get-Date
        ErrorMessage     = $null
    }

    try {
        # One-time backup BEFORE changes
        $backupPath = "$path.bak"
        if (-not (Test-Path $backupPath)) {
            Copy-Item -LiteralPath $path -Destination $backupPath
            $entry.BackupCreated = $true
        }

        # Read file (detect encoding + BOM)
        $file = Read-TextWithEncodingDetection -Path $path
        $text = $file.Text

        # Step 1: patch only the tag text (no XML save)
        $tag = Find-FirstActiveDataSettingsTag -Text $text -LookAheadChars $LookAheadChars
        if (-not $tag) {
            $entry.PatchStatus = 'MissingDataSettings'
        }
        else {
            $tagText = $text.Substring($tag.Start, $tag.End - $tag.Start + 1)
            $patch = Patch-DataSettingsTag -TagText $tagText -DesiredAttributes $desiredAttributes

            if ($patch.Changed) {
                $newText = $text.Substring(0, $tag.Start) + $patch.NewTag + $text.Substring($tag.End + 1)

                # Write back with ORIGINAL encoding/BOM to avoid any non-target changes
                Write-TextPreserveEncoding -Path $path -Text $newText -Encoding $file.Encoding -HasBom $file.HasBom

                $entry.PatchStatus = 'Patched'
            }
            else {
                $entry.PatchStatus = 'AlreadyCompliant'
            }
        }

        # Step 2: encoding conversion (optional) — only if BOM exists
        if ($ConvertToUtf8NoBom_OnlyIfBomExists) {
            $conv = Convert-ToUtf8NoBom-IfBom -Path $path
            if ($conv.Converted) {
                $entry.EncodingStatus = 'Converted_ToUtf8NoBOM'
                $entry.BomType = $conv.BomType
                $entry.FromEncoding = $conv.FromEnc
            }
            else {
                $entry.EncodingStatus = 'Skipped_NoBOM'
            }
        }
        else {
            $entry.EncodingStatus = 'NotRequested'
        }
    }
    catch {
        $entry.PatchStatus = 'Error'
        $entry.EncodingStatus = 'Error'
        $entry.ErrorMessage = $_.Exception.Message
    }

    $log.Add($entry)
}

# ============================
# OUTPUT LOG + SUMMARY
# ============================

$log | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

$patched     = @($log | Where-Object PatchStatus -eq 'Patched').Count
$compliant   = @($log | Where-Object PatchStatus -eq 'AlreadyCompliant').Count
$missing     = @($log | Where-Object PatchStatus -eq 'MissingDataSettings').Count
$converted   = @($log | Where-Object EncodingStatus -eq 'Converted_ToUtf8NoBOM').Count
$skipped     = @($log | Where-Object EncodingStatus -eq 'Skipped_NoBOM').Count
$errors      = @($log | Where-Object { $_.PatchStatus -eq 'Error' -or $_.EncodingStatus -eq 'Error' }).Count

Write-Host "Completed." -ForegroundColor Cyan
Write-Host ("  Patched:              {0}" -f $patched)
Write-Host ("  AlreadyCompliant:     {0}" -f $compliant)
Write-Host ("  MissingDataSettings:  {0}" -f $missing)
Write-Host ("  Converted (BOM->UTF8 no BOM): {0}" -f $converted)
Write-Host ("  Skipped (no BOM):     {0}" -f $skipped)
Write-Host ("  Errors:               {0}" -f $errors) -ForegroundColor ($(if($errors){'Red'}else{'Gray'}))
Write-Host "Log written to: $csvPath" -ForegroundColor Green
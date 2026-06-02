<#
.SYNOPSIS
Updates columns 7 and 8 in two fixed-width text files.

.DESCRIPTION
Uses the fixed-width layouts below:

1st file: 13, 16, 26, 12, 8, 16, 10, 12
2nd file: 26, 13, 9, 12, 8, 26, 10, 12

The same two values are applied to columns 7 and 8 in both files. Values are
padded with spaces to fit each file's required column width. If a value is too
long for either file, the script stops unless -TruncateValues is supplied.

For the first file only, the script repairs rows longer than the expected 113
characters by finding the " - " marker that starts column 4 and removing the
extra spaces immediately before it. All first-file repairs are completed before
columns 7 and 8 are updated.
If a first-file row cannot be safely repaired, the script stops before writing
output unless -AllowUnrepairableFirstRows is supplied.

.EXAMPLE
.\Update-FixedWidthColumns.ps1 `
  -FirstFilePath  "C:\Data\1st.txt" `
  -SecondFilePath "C:\Data\2nd.txt" `
  -Column7Value "ABC123" `
  -Column8Value "UPDATED"

Writes C:\Data\1st.updated.txt and C:\Data\2nd.updated.txt.

.EXAMPLE
.\Update-FixedWidthColumns.ps1 `
  -FirstFilePath  "C:\Data\1st.txt" `
  -SecondFilePath "C:\Data\2nd.txt" `
  -Column7Value "ABC123" `
  -Column8Value "UPDATED" `
  -InPlace `
  -CreateBackup

Updates the original files and creates .bak copies first.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$FirstFilePath,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$SecondFilePath,

    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Column7Value,

    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Column8Value,

    [switch]$InPlace,

    [string]$FirstOutputPath,

    [string]$SecondOutputPath,

    [switch]$TruncateValues,

    [switch]$AllowUnrepairableFirstRows,

    [switch]$CreateBackup,

    [ValidateSet("UTF8", "UTF8BOM", "ASCII", "Unicode", "BigEndianUnicode", "UTF32", "Default", "OEM")]
    [string]$Encoding = "UTF8"
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$FirstLayout = @{
    Name          = "1st file"
    Widths        = @(13, 16, 26, 12, 8, 16, 10, 12)
    RepairColumn3 = $true
}

$SecondLayout = @{
    Name          = "2nd file"
    Widths        = @(26, 13, 9, 12, 8, 26, 10, 12)
    RepairColumn3 = $false
}

function Get-TextEncoding {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    switch ($Name) {
        "UTF8"             { return New-Object System.Text.UTF8Encoding -ArgumentList $false }
        "UTF8BOM"          { return New-Object System.Text.UTF8Encoding -ArgumentList $true }
        "ASCII"            { return [System.Text.Encoding]::ASCII }
        "Unicode"          { return [System.Text.Encoding]::Unicode }
        "BigEndianUnicode" { return [System.Text.Encoding]::BigEndianUnicode }
        "UTF32"            { return [System.Text.Encoding]::UTF32 }
        "Default"          { return [System.Text.Encoding]::Default }
        "OEM" {
            $codePage = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage
            return [System.Text.Encoding]::GetEncoding($codePage)
        }
    }
}

function Get-FullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function Get-UpdatedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = Get-FullPath -Path $Path
    $directory = Split-Path -Parent $fullPath
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
    $extension = [System.IO.Path]::GetExtension($fullPath)

    return Join-Path -Path $directory -ChildPath "$fileName.updated$extension"
}

function Get-TotalWidth {
    param(
        [Parameter(Mandatory = $true)]
        [int[]]$Widths
    )

    $total = 0
    foreach ($width in $Widths) {
        $total += $width
    }

    return $total
}

function Get-ColumnStart {
    param(
        [Parameter(Mandatory = $true)]
        [int[]]$Widths,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 8)]
        [int]$Column
    )

    $start = 0
    for ($index = 0; $index -lt ($Column - 1); $index++) {
        $start += $Widths[$index]
    }

    return $start
}

function Format-FixedWidthValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [int]$Width,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ($Value.Length -gt $Width) {
        if ($TruncateValues) {
            return $Value.Substring(0, $Width)
        }

        throw "$Description is $($Value.Length) characters, but the column width is $Width. Shorten the value or rerun with -TruncateValues."
    }

    return $Value.PadRight($Width, " ")
}

function Repair-FirstFileColumn3 {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Line,

        [Parameter(Mandatory = $true)]
        [int[]]$Widths,

        [Parameter(Mandatory = $true)]
        [int]$LineNumber
    )

    $totalWidth = Get-TotalWidth -Widths $Widths
    $column4Start = Get-ColumnStart -Widths $Widths -Column 4

    if ($Line.Length -le $totalWidth) {
        return [pscustomobject]@{
            Line     = $Line
            Repaired = $false
        }
    }

    $excessLength = $Line.Length - $totalWidth
    $candidatePositions = New-Object System.Collections.Generic.List[int]

    for ($index = 1; $index -lt ($Line.Length - 1); $index++) {
        if ($Line[$index] -eq "-" -and [char]::IsWhiteSpace($Line[$index - 1]) -and [char]::IsWhiteSpace($Line[$index + 1])) {
            $candidatePositions.Add($index)
        }
    }

    $actualColumn4Start = $null
    foreach ($candidatePosition in $candidatePositions) {
        if ($candidatePosition -le $column4Start) {
            continue
        }

        $charactersToRemove = $Line.Substring($column4Start, $candidatePosition - $column4Start)
        if ($charactersToRemove -match "^\s+$") {
            $actualColumn4Start = $candidatePosition
            break
        }
    }

    if ($null -eq $actualColumn4Start) {
        $candidateDisplay = if ($candidatePositions.Count -gt 0) {
            (($candidatePositions | ForEach-Object { $_ + 1 }) -join ", ")
        }
        else {
            "none"
        }

        $message = "1st file line ${LineNumber}: row is $excessLength characters too long, but no usable whitespace-hyphen-whitespace column 4 marker was found after position $($column4Start + 1). Candidate marker positions: $candidateDisplay."
        if (-not $AllowUnrepairableFirstRows) {
            throw "$message The script stopped before updating columns 7 and 8."
        }

        Write-Warning "$message The row was left at its current length before updating columns 7 and 8."
        return [pscustomobject]@{
            Line     = $Line
            Repaired = $false
        }
    }

    $charactersToRemove = $Line.Substring($column4Start, $actualColumn4Start - $column4Start)
    if ($charactersToRemove -notmatch "^\s+$") {
        $message = "1st file line ${LineNumber}: row is $excessLength characters too long, but the characters before the column 4 marker are not all whitespace."
        if (-not $AllowUnrepairableFirstRows) {
            throw "$message The script stopped before updating columns 7 and 8."
        }

        Write-Warning "$message The row was left at its current length before updating columns 7 and 8."
        return [pscustomobject]@{
            Line     = $Line
            Repaired = $false
        }
    }

    $repairedLine = $Line.Substring(0, $column4Start) + $Line.Substring($actualColumn4Start)
    if ($repairedLine.Length -ne $totalWidth) {
        $message = "1st file line ${LineNumber}: repaired row length is $($repairedLine.Length), but expected $totalWidth."
        if (-not $AllowUnrepairableFirstRows) {
            throw "$message The script stopped before updating columns 7 and 8."
        }

        Write-Warning "$message The row was repaired and columns 7 and 8 will be updated at the expected positions."
    }

    return [pscustomobject]@{
        Line     = $repairedLine
        Repaired = $true
    }
}

function Set-FixedWidthColumns {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Line,

        [Parameter(Mandatory = $true)]
        [hashtable]$Layout,

        [Parameter(Mandatory = $true)]
        [int]$LineNumber
    )

    if ($Line.Length -eq 0) {
        return $Line
    }

    $widths = [int[]]$Layout.Widths
    $totalWidth = Get-TotalWidth -Widths $widths
    $column7Start = Get-ColumnStart -Widths $widths -Column 7
    $column8Start = Get-ColumnStart -Widths $widths -Column 8
    $column7Width = $widths[6]
    $column8Width = $widths[7]

    if ($Line.Length -lt $totalWidth) {
        $Line = $Line.PadRight($totalWidth, " ")
    }
    elseif ($Line.Length -gt $totalWidth) {
        Write-Warning "$($Layout.Name) line $LineNumber is $($Line.Length) characters after repair; expected $totalWidth. Columns 7 and 8 were updated at the expected positions, but the row still contains extra characters."
    }

    $formattedColumn7 = Format-FixedWidthValue -Value $Column7Value -Width $column7Width -Description "$($Layout.Name) column 7 value"
    $formattedColumn8 = Format-FixedWidthValue -Value $Column8Value -Width $column8Width -Description "$($Layout.Name) column 8 value"

    $Line = $Line.Substring(0, $column7Start) +
        $formattedColumn7 +
        $Line.Substring($column7Start + $column7Width)

    $Line = $Line.Substring(0, $column8Start) +
        $formattedColumn8 +
        $Line.Substring($column8Start + $column8Width)

    return $Line
}

function Repair-FixedWidthLines {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Lines,

        [Parameter(Mandatory = $true)]
        [hashtable]$Layout
    )

    $repairedLines = New-Object System.Collections.Generic.List[string]
    $repairedCount = 0

    for ($index = 0; $index -lt $Lines.Count; $index++) {
        $lineNumber = $index + 1
        $line = $Lines[$index]

        if ($Layout.RepairColumn3) {
            $repairResult = Repair-FirstFileColumn3 -Line $line -Widths ([int[]]$Layout.Widths) -LineNumber $lineNumber
            $line = $repairResult.Line

            if ($repairResult.Repaired) {
                $repairedCount++
            }
        }

        $repairedLines.Add($line)
    }

    return [pscustomobject]@{
        Lines        = [string[]]$repairedLines.ToArray()
        RepairedRows = $repairedCount
    }
}

function Update-FixedWidthLines {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Lines,

        [Parameter(Mandatory = $true)]
        [hashtable]$Layout
    )

    $updatedLines = New-Object System.Collections.Generic.List[string]

    for ($index = 0; $index -lt $Lines.Count; $index++) {
        $lineNumber = $index + 1
        $updatedLines.Add((Set-FixedWidthColumns -Line $Lines[$index] -Layout $Layout -LineNumber $lineNumber))
    }

    return [string[]]$updatedLines.ToArray()
}

function Update-FixedWidthFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [hashtable]$Layout,

        [Parameter(Mandatory = $true)]
        [System.Text.Encoding]$TextEncoding
    )

    $inputFullPath = (Resolve-Path -LiteralPath $InputPath).Path
    $outputFullPath = Get-FullPath -Path $OutputPath
    $outputDirectory = Split-Path -Parent $outputFullPath

    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }

    $lines = [System.IO.File]::ReadAllLines($inputFullPath, $TextEncoding)
    $repairResult = Repair-FixedWidthLines -Lines $lines -Layout $Layout
    $updatedLines = Update-FixedWidthLines -Lines ([string[]]$repairResult.Lines) -Layout $Layout

    if ($PSCmdlet.ShouldProcess($outputFullPath, "write updated fixed-width file")) {
        if ($CreateBackup -and ($inputFullPath -eq $outputFullPath)) {
            Copy-Item -LiteralPath $inputFullPath -Destination "$inputFullPath.bak" -Force
        }

        [System.IO.File]::WriteAllLines($outputFullPath, $updatedLines, $TextEncoding)
    }

    return [pscustomobject]@{
        File          = $outputFullPath
        Rows          = $lines.Count
        RepairedRows  = $repairResult.RepairedRows
        UpdatedInFile = $inputFullPath -eq $outputFullPath
    }
}

if ($InPlace -and ($FirstOutputPath -or $SecondOutputPath)) {
    throw "Use either -InPlace or output paths, not both."
}

if ($InPlace) {
    $FirstOutputPath = $FirstFilePath
    $SecondOutputPath = $SecondFilePath
}
else {
    if (-not $FirstOutputPath) {
        $FirstOutputPath = Get-UpdatedPath -Path $FirstFilePath
    }

    if (-not $SecondOutputPath) {
        $SecondOutputPath = Get-UpdatedPath -Path $SecondFilePath
    }
}

$textEncoding = Get-TextEncoding -Name $Encoding

$firstResult = Update-FixedWidthFile -InputPath $FirstFilePath -OutputPath $FirstOutputPath -Layout $FirstLayout -TextEncoding $textEncoding
$secondResult = Update-FixedWidthFile -InputPath $SecondFilePath -OutputPath $SecondOutputPath -Layout $SecondLayout -TextEncoding $textEncoding

$firstResult
$secondResult

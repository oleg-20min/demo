$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

$inputPath = Join-Path $PSScriptRoot "bfs_domestic_violence_2009_2024.xlsx"
$outputPath = Join-Path $PSScriptRoot "bfs_domestic_violence_flat.csv"

function Get-ZipEntryText {
    param(
        [Parameter(Mandatory = $true)] $Zip,
        [Parameter(Mandatory = $true)] [string] $EntryPath
    )

    $entry = $Zip.Entries | Where-Object { $_.FullName -eq $EntryPath }
    if (-not $entry) {
        return $null
    }

    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-XmlDocument {
    param(
        [Parameter(Mandatory = $true)] $Zip,
        [Parameter(Mandatory = $true)] [string] $EntryPath
    )

    $text = Get-ZipEntryText -Zip $Zip -EntryPath $EntryPath
    if ($null -eq $text) {
        return $null
    }

    return [xml]$text
}

function Get-ColumnIndex {
    param(
        [Parameter(Mandatory = $true)] [string] $CellReference
    )

    $letters = ($CellReference -replace "\d", "").ToUpperInvariant()
    $index = 0
    foreach ($char in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
    }

    return $index
}

function Get-SharedStrings {
    param(
        [Parameter(Mandatory = $true)] $Zip
    )

    $xml = Get-XmlDocument -Zip $Zip -EntryPath "xl/sharedStrings.xml"
    if ($null -eq $xml) {
        return @()
    }

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    $result = New-Object System.Collections.Generic.List[string]
    foreach ($si in $xml.SelectNodes("//d:sst/d:si", $ns)) {
        $parts = $si.SelectNodes(".//d:t", $ns) | ForEach-Object { $_.'#text' }
        $result.Add(($parts -join ""))
    }

    return $result.ToArray()
}

function Get-CellStringValue {
    param(
        [Parameter(Mandatory = $true)] $CellNode,
        [Parameter(Mandatory = $true)] [string[]] $SharedStrings,
        [Parameter(Mandatory = $true)] $NamespaceManager
    )

    $type = [string]$CellNode.t
    if ($type -eq "inlineStr") {
        $parts = $CellNode.SelectNodes(".//d:is/d:t", $NamespaceManager) | ForEach-Object { $_.'#text' }
        return ($parts -join "")
    }

    $valueNode = $CellNode.SelectSingleNode("./d:v", $NamespaceManager)
    if ($null -eq $valueNode) {
        return ""
    }

    $raw = [string]$valueNode.InnerText
    if ($type -eq "s") {
        $index = [int]$raw
        if ($index -ge 0 -and $index -lt $SharedStrings.Length) {
            return $SharedStrings[$index]
        }
        return ""
    }

    return $raw
}

function Convert-BfsValue {
    param(
        [AllowNull()][string] $Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $normalized = $Text.Replace([char]0xA0, " ").Trim()
    if ($normalized -eq "X") {
        return $null
    }

    $normalized = $normalized -replace "[,\s']", ""
    if ($normalized -match "^-?\d+(\.\d+)?$") {
        return [double]::Parse($normalized, [System.Globalization.CultureInfo]::InvariantCulture)
    }

    return $null
}

function Is-RelationshipRow {
    param(
        [AllowNull()][string] $RelationshipText
    )

    if ([string]::IsNullOrWhiteSpace($RelationshipText)) {
        return $false
    }

    return $RelationshipText -in @(
        "Total",
        "partnership",
        "former partnership",
        "parent-child relationship",
        "other family relationship"
    )
}

function Get-SheetMap {
    param(
        [Parameter(Mandatory = $true)] $Zip
    )

    $workbookXml = Get-XmlDocument -Zip $Zip -EntryPath "xl/workbook.xml"
    $relsXml = Get-XmlDocument -Zip $Zip -EntryPath "xl/_rels/workbook.xml.rels"

    $wbNs = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
    $wbNs.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $relNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
    $relNs.AddNamespace("d", "http://schemas.openxmlformats.org/package/2006/relationships")

    $targetsById = @{}
    foreach ($rel in $relsXml.SelectNodes("//d:Relationships/d:Relationship", $relNs)) {
        $targetsById[[string]$rel.Id] = [string]$rel.Target
    }

    $sheets = New-Object System.Collections.Generic.List[object]
    foreach ($sheet in $workbookXml.SelectNodes("//d:sheets/d:sheet", $wbNs)) {
        $relationshipId = [string]$sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $target = $targetsById[$relationshipId]
        if (-not $target) {
            continue
        }

        $normalizedTarget = if ($target.StartsWith("/")) { $target.TrimStart("/") } else { "xl/$target" }
        $sheets.Add([pscustomobject]@{
            Name = [string]$sheet.name
            Path = $normalizedTarget
        })
    }

    return $sheets
}

$sharedStrings = @()
$records = New-Object System.Collections.Generic.List[object]
$zip = [System.IO.Compression.ZipFile]::OpenRead($inputPath)

try {
    $sharedStrings = Get-SharedStrings -Zip $zip
    $sheetMap = Get-SheetMap -Zip $zip

    foreach ($sheet in $sheetMap) {
        if ($sheet.Name -notmatch "^\d{4}$") {
            continue
        }

        $sheetXml = Get-XmlDocument -Zip $zip -EntryPath $sheet.Path
        if ($null -eq $sheetXml) {
            continue
        }

        $ns = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
        $ns.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        $currentOffenceType = $null
        $currentTotalOffences = $null
        $currentGender = $null

        foreach ($rowNode in $sheetXml.SelectNodes("//d:worksheet/d:sheetData/d:row", $ns)) {
            $rowNumber = [int]$rowNode.r
            if ($rowNumber -lt 7) {
                continue
            }

            $cells = @{}
            foreach ($cellNode in $rowNode.SelectNodes("./d:c", $ns)) {
                $columnIndex = Get-ColumnIndex -CellReference ([string]$cellNode.r)
                if ($columnIndex -in @(1, 2, 4, 7)) {
                    $cells[$columnIndex] = (Get-CellStringValue -CellNode $cellNode -SharedStrings $sharedStrings -NamespaceManager $ns).Trim()
                }
            }

            $col1 = if ($cells.ContainsKey(1)) { $cells[1] } else { "" }
            $col2 = if ($cells.ContainsKey(2)) { $cells[2] } else { "" }
            $col4 = if ($cells.ContainsKey(4)) { $cells[4] } else { "" }
            $col7 = if ($cells.ContainsKey(7)) { $cells[7] } else { "" }

            if ($col1 -like "X:*" -or $col1 -like "1)*" -or $col1 -like "State of the database:*" -or $col1 -like "Source:*" -or $col1 -like "Information:*") {
                break
            }

            if ([string]::IsNullOrWhiteSpace($col1) -and [string]::IsNullOrWhiteSpace($col2)) {
                continue
            }

            $normalizedRelationship = $col2 -replace "^\s+", ""

            if (-not [string]::IsNullOrWhiteSpace($col1) -and $col1 -notin @("male", "female")) {
                $currentOffenceType = $col1
                $currentTotalOffences = Convert-BfsValue $col4
                $currentGender = $null
                continue
            }

            if ($col1 -in @("male", "female")) {
                $currentGender = $col1
                $relationshipType = if (Is-RelationshipRow $normalizedRelationship) { $normalizedRelationship } else { "Total" }
                $clearedOffences = Convert-BfsValue $col7
            }
            elseif ([string]::IsNullOrWhiteSpace($col1) -and $currentGender -and (Is-RelationshipRow $normalizedRelationship)) {
                $relationshipType = $normalizedRelationship
                $clearedOffences = Convert-BfsValue $col7
            }
            else {
                continue
            }

            $clearanceRate = $null
            if ($null -ne $currentTotalOffences -and $currentTotalOffences -ne 0 -and $null -ne $clearedOffences) {
                $clearanceRate = [math]::Round($clearedOffences / $currentTotalOffences, 6)
            }

            $records.Add([pscustomobject]@{
                Year               = [int]$sheet.Name
                Offence_Type       = $currentOffenceType
                Gender_Accused     = $currentGender
                Relationship_Type  = $relationshipType
                Total_Offences     = $currentTotalOffences
                Cleared_Offences   = $clearedOffences
                Clearance_Rate     = $clearanceRate
            })
        }
    }
}
finally {
    $zip.Dispose()
}

$records |
    Sort-Object Year, Offence_Type, Gender_Accused, Relationship_Type |
    Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Output "Wrote $($records.Count) rows to $outputPath"

[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$InputJson,

    [string]$TemplateXlsx,

    [string]$OutputXlsx,

    [string]$NameNo,

    [string]$Name,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RemainingArguments
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$positionalArguments = [System.Collections.Generic.List[string]]::new()
$rawArguments = @($RemainingArguments)
for ($argumentIndex = 0; $argumentIndex -lt $rawArguments.Count; $argumentIndex++) {
    $argument = $rawArguments[$argumentIndex]
    switch ($argument.ToLowerInvariant()) {
        '--nameno' {
            if (($argumentIndex + 1) -ge $rawArguments.Count) {
                throw 'The --nameno option requires a value.'
            }

            $argumentIndex++
            $NameNo = $rawArguments[$argumentIndex]
            continue
        }
        '--name' {
            if (($argumentIndex + 1) -ge $rawArguments.Count) {
                throw 'The --name option requires a value.'
            }

            $argumentIndex++
            $Name = $rawArguments[$argumentIndex]
            continue
        }
        default {
            [void]$positionalArguments.Add($argument)
        }
    }
}

foreach ($argument in $positionalArguments) {
    if ([string]::IsNullOrWhiteSpace($InputJson)) {
        $InputJson = $argument
    }
    elseif ([string]::IsNullOrWhiteSpace($TemplateXlsx)) {
        $TemplateXlsx = $argument
    }
    elseif ([string]::IsNullOrWhiteSpace($OutputXlsx)) {
        $OutputXlsx = $argument
    }
    else {
        throw "Unexpected positional argument: $argument"
    }
}

if ([string]::IsNullOrWhiteSpace($InputJson)) {
    throw 'InputJson is required.'
}

if ([string]::IsNullOrWhiteSpace($TemplateXlsx)) {
    $TemplateXlsx = 'seikatsu.xlsx'
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web.Extensions

$script:SpreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:OfficeDocumentRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:PackageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
$script:ContentTypesNamespace = 'http://schemas.openxmlformats.org/package/2006/content-types'
$script:AppPropertiesNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
$script:VariantTypesNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
$script:CorePropertiesNamespace = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
$script:DublinCoreNamespace = 'http://purl.org/dc/elements/1.1/'
$script:DublinCoreTermsNamespace = 'http://purl.org/dc/terms/'
$script:SpreadsheetRevisionNamespace = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision'
$script:XmlNamespace = 'http://www.w3.org/XML/1998/namespace'

$script:DayStartRows = @(14..79 | Where-Object { (($_ - 14) % 5) -eq 0 })
$script:PrintableArea = '$A$1:$BG$83'
$script:DimensionRef = 'A1:BH128'
$script:NoteColumnNumber = 60
$script:TimeSlotStartColumn = 9
$script:HalfHourSlotsPerDay = 48
$script:PcOffColumnNumber = 44
$script:SleepTotalColumnNumber = 58
$script:PcOffText = '22' + ([string][char]0x6642) + ([string][char]0x306B) + 'PC' + ([string][char]0x30AA) + ([string][char]0x30D5)
$script:SleepActivityType = ([string][char]0x7761) + ([string][char]0x7720)

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$MustExist
    )

    $absolutePath = if ([System.IO.Path]::IsPathRooted($Path)) {
        [System.IO.Path]::GetFullPath($Path)
    }
    else {
        [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location) -ChildPath $Path))
    }

    if ($MustExist -and -not (Test-Path -LiteralPath $absolutePath)) {
        throw "Path not found: $absolutePath"
    }

    return $absolutePath
}

function ConvertTo-ColumnName {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ColumnNumber
    )

    if ($ColumnNumber -lt 1) {
        throw "ColumnNumber must be 1 or greater: $ColumnNumber"
    }

    $name = ''
    $index = $ColumnNumber
    while ($index -gt 0) {
        $index--
        $name = ([char][int](65 + ($index % 26))) + $name
        $index = [math]::Floor($index / 26)
    }

    return $name
}

function ConvertTo-ColumnNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ColumnName
    )

    $value = 0
    foreach ($character in $ColumnName.ToUpperInvariant().ToCharArray()) {
        $value = ($value * 26) + ([int][char]$character - [int][char]'A' + 1)
    }

    return $value
}

function Get-ColumnNumberFromCellReference {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CellReference
    )

    $match = [regex]::Match($CellReference, '^[A-Z]+')
    if (-not $match.Success) {
        throw "Invalid cell reference: $CellReference"
    }

    return ConvertTo-ColumnNumber -ColumnName $match.Value
}

function New-NamespaceManager {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Document
    )

    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
    $namespaceManager.AddNamespace('x', $script:SpreadsheetNamespace)
    $namespaceManager.AddNamespace('r', $script:OfficeDocumentRelationshipNamespace)
    $namespaceManager.AddNamespace('pr', $script:PackageRelationshipNamespace)
    $namespaceManager.AddNamespace('ct', $script:ContentTypesNamespace)
    $namespaceManager.AddNamespace('ap', $script:AppPropertiesNamespace)
    $namespaceManager.AddNamespace('vt', $script:VariantTypesNamespace)
    $namespaceManager.AddNamespace('cp', $script:CorePropertiesNamespace)
    $namespaceManager.AddNamespace('dc', $script:DublinCoreNamespace)
    $namespaceManager.AddNamespace('dcterms', $script:DublinCoreTermsNamespace)
    $namespaceManager.AddNamespace('xr', $script:SpreadsheetRevisionNamespace)

    return ,$namespaceManager
}

function Load-XmlDocument {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $document = New-Object System.Xml.XmlDocument
    $document.PreserveWhitespace = $true
    $document.Load($Path)
    return $document
}

function Write-ZipArchiveFromDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$DestinationFile
    )

    if (Test-Path -LiteralPath $DestinationFile) {
        Remove-Item -LiteralPath $DestinationFile -Force
    }

    $archive = [System.IO.Compression.ZipFile]::Open($DestinationFile, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $rootPath = (Resolve-AbsolutePath -Path $SourceDirectory -MustExist).TrimEnd('\')
        foreach ($file in Get-ChildItem -LiteralPath $rootPath -Recurse -File) {
            $relativePath = $file.FullName.Substring($rootPath.Length + 1).Replace('\', '/')
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive,
                $file.FullName,
                $relativePath,
                [System.IO.Compression.CompressionLevel]::Optimal
            )
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Set-NamespacedAttribute {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Element,

        [Parameter(Mandatory = $true)]
        [string]$Prefix,

        [Parameter(Mandatory = $true)]
        [string]$LocalName,

        [Parameter(Mandatory = $true)]
        [string]$NamespaceUri,

        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $existingAttribute = $Element.Attributes.GetNamedItem($LocalName, $NamespaceUri)
    if ($existingAttribute) {
        $existingAttribute.Value = $Value
        return
    }

    $attribute = $Element.OwnerDocument.CreateAttribute($Prefix, $LocalName, $NamespaceUri)
    $attribute.Value = $Value
    [void]$Element.Attributes.Append($attribute)
}

function Get-OrCreate-CellNode {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$RowNode,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager,

        [Parameter(Mandatory = $true)]
        [string]$CellReference,

        [string]$StyleIndex
    )

    $cellNode = $RowNode.SelectSingleNode("./x:c[@r='$CellReference']", $NamespaceManager)
    if ($cellNode -is [System.Xml.XmlElement]) {
        if ($StyleIndex) {
            $cellNode.SetAttribute('s', $StyleIndex)
        }
        return $cellNode
    }

    $cellNode = $RowNode.OwnerDocument.CreateElement('c', $script:SpreadsheetNamespace)
    $cellNode.SetAttribute('r', $CellReference)
    if ($StyleIndex) {
        $cellNode.SetAttribute('s', $StyleIndex)
    }

    $targetColumn = Get-ColumnNumberFromCellReference -CellReference $CellReference
    $inserted = $false
    foreach ($existingCell in @($RowNode.SelectNodes('./x:c', $NamespaceManager))) {
        $existingColumn = Get-ColumnNumberFromCellReference -CellReference $existingCell.GetAttribute('r')
        if ($existingColumn -gt $targetColumn) {
            [void]$RowNode.InsertBefore($cellNode, $existingCell)
            $inserted = $true
            break
        }
    }

    if (-not $inserted) {
        [void]$RowNode.AppendChild($cellNode)
    }

    return $cellNode
}

function Clear-CellValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$CellNode
    )

    $CellNode.RemoveAttribute('t')
    while ($CellNode.HasChildNodes) {
        [void]$CellNode.RemoveChild($CellNode.FirstChild)
    }
}

function Set-CellNumber {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$RowNode,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager,

        [Parameter(Mandatory = $true)]
        [string]$CellReference,

        [Parameter(Mandatory = $true)]
        [int]$Value,

        [string]$StyleIndex
    )

    $cellNode = Get-OrCreate-CellNode -RowNode $RowNode -NamespaceManager $NamespaceManager -CellReference $CellReference -StyleIndex $StyleIndex
    Clear-CellValue -CellNode $cellNode

    $valueNode = $RowNode.OwnerDocument.CreateElement('v', $script:SpreadsheetNamespace)
    $valueNode.InnerText = [string]$Value
    [void]$cellNode.AppendChild($valueNode)
}

function Set-CellDecimalNumber {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$RowNode,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager,

        [Parameter(Mandatory = $true)]
        [string]$CellReference,

        [Parameter(Mandatory = $true)]
        [double]$Value,

        [string]$StyleIndex
    )

    $cellNode = Get-OrCreate-CellNode -RowNode $RowNode -NamespaceManager $NamespaceManager -CellReference $CellReference -StyleIndex $StyleIndex
    Clear-CellValue -CellNode $cellNode

    $valueNode = $RowNode.OwnerDocument.CreateElement('v', $script:SpreadsheetNamespace)
    $valueNode.InnerText = $Value.ToString('0.###', [System.Globalization.CultureInfo]::InvariantCulture)
    [void]$cellNode.AppendChild($valueNode)
}

function Set-CellInlineString {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$RowNode,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager,

        [Parameter(Mandatory = $true)]
        [string]$CellReference,

        [Parameter(Mandatory = $true)]
        [string]$Value,

        [string]$StyleIndex
    )

    $cellNode = Get-OrCreate-CellNode -RowNode $RowNode -NamespaceManager $NamespaceManager -CellReference $CellReference -StyleIndex $StyleIndex
    Clear-CellValue -CellNode $cellNode
    $cellNode.SetAttribute('t', 'inlineStr')

    $inlineStringNode = $RowNode.OwnerDocument.CreateElement('is', $script:SpreadsheetNamespace)
    $textNode = $RowNode.OwnerDocument.CreateElement('t', $script:SpreadsheetNamespace)
    Set-NamespacedAttribute -Element $textNode -Prefix 'xml' -LocalName 'space' -NamespaceUri $script:XmlNamespace -Value 'preserve'
    $textNode.InnerText = $Value
    [void]$inlineStringNode.AppendChild($textNode)
    [void]$cellNode.AppendChild($inlineStringNode)
}

function Set-CellStyle {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$RowNode,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager]$NamespaceManager,

        [Parameter(Mandatory = $true)]
        [string]$CellReference,

        [Parameter(Mandatory = $true)]
        [string]$StyleIndex
    )

    $cellNode = Get-OrCreate-CellNode -RowNode $RowNode -NamespaceManager $NamespaceManager -CellReference $CellReference
    $cellNode.SetAttribute('s', $StyleIndex)
}

function Set-RowMaxSpan {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$RowNode,

        [Parameter(Mandatory = $true)]
        [int]$MaxColumnNumber
    )

    $currentSpans = $RowNode.GetAttribute('spans')
    if ([string]::IsNullOrWhiteSpace($currentSpans)) {
        $RowNode.SetAttribute('spans', "1:$MaxColumnNumber")
        return
    }

    $parts = $currentSpans.Split(':')
    if ($parts.Count -ne 2) {
        $RowNode.SetAttribute('spans', "1:$MaxColumnNumber")
        return
    }

    $minSpan = [int]$parts[0]
    $currentMax = [int]$parts[1]
    if ($currentMax -lt $MaxColumnNumber) {
        $RowNode.SetAttribute('spans', ('{0}:{1}' -f $minSpan, $MaxColumnNumber))
    }
}

function Get-DayOfWeekLabel {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$Date
    )

    $labels = @{
        Monday    = ([string][char]0x6708)
        Tuesday   = ([string][char]0x706B)
        Wednesday = ([string][char]0x6C34)
        Thursday  = ([string][char]0x6728)
        Friday    = ([string][char]0x91D1)
        Saturday  = ([string][char]0x571F)
        Sunday    = ([string][char]0x65E5)
    }

    $key = $Date.DayOfWeek.ToString()
    if (-not $labels.ContainsKey($key)) {
        throw "Unsupported DayOfWeek value: $($Date.DayOfWeek)"
    }

    return $labels[$key]
}
function Floor-ToHalfHour {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$DateTime
    )

    $roundedMinutes = [math]::Floor((($DateTime - $DateTime.Date).TotalMinutes) / 30) * 30
    return $DateTime.Date.AddMinutes($roundedMinutes)
}

function Ceiling-ToHalfHour {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$DateTime
    )

    $roundedMinutes = [math]::Ceiling((($DateTime - $DateTime.Date).TotalMinutes) / 30) * 30
    return $DateTime.Date.AddMinutes($roundedMinutes)
}

function Add-DaySegmentToSlots {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$DayState,

        [Parameter(Mandatory = $true)]
        [datetime]$SegmentStart,

        [Parameter(Mandatory = $true)]
        [datetime]$SegmentEnd,

        [Parameter(Mandatory = $true)]
        [string]$ActivityType
    )

    if ($SegmentEnd -le $SegmentStart) {
        return
    }

    $nextMidnight = $SegmentStart.Date.AddDays(1)
    if ($SegmentEnd -gt $nextMidnight) {
        Add-DaySegmentToSlots -DayState $DayState -SegmentStart $SegmentStart -SegmentEnd $nextMidnight -ActivityType $ActivityType
        Add-DaySegmentToSlots -DayState $DayState -SegmentStart $nextMidnight -SegmentEnd $SegmentEnd -ActivityType $ActivityType
        return
    }

    $dayBoundary = $SegmentStart.Date.AddHours(4)
    if ($SegmentStart -lt $dayBoundary -and $SegmentEnd -gt $dayBoundary) {
        Add-DaySegmentToSlots -DayState $DayState -SegmentStart $SegmentStart -SegmentEnd $dayBoundary -ActivityType $ActivityType
        Add-DaySegmentToSlots -DayState $DayState -SegmentStart $dayBoundary -SegmentEnd $SegmentEnd -ActivityType $ActivityType
        return
    }

    $slots = $DayState['Slots']
    $segmentStartMinutes = [int](($SegmentStart - $SegmentStart.Date).TotalMinutes)
    $segmentEndMinutes = if ($SegmentEnd.Date -gt $SegmentStart.Date) {
        1440
    }
    else {
        [int](($SegmentEnd - $SegmentEnd.Date).TotalMinutes)
    }

    if ($segmentEndMinutes -le 240) {
        $slotStart = 40 + [int]($segmentStartMinutes / 30)
        $slotEndExclusive = 40 + [int]($segmentEndMinutes / 30)
    }
    else {
        $slotStart = [int](($segmentStartMinutes - 240) / 30)
        $slotEndExclusive = [int](($segmentEndMinutes - 240) / 30)
    }

    for ($slotIndex = $slotStart; $slotIndex -lt $slotEndExclusive; $slotIndex++) {
        if ($slotIndex -lt 0 -or $slotIndex -ge $script:HalfHourSlotsPerDay) {
            throw "Calculated slot index out of range: $slotIndex"
        }

        $Slots[$slotIndex] = $ActivityType
    }
}

function Add-ActivityToDayMap {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$DayMap,

        [Parameter(Mandatory = $true)]
        [datetime]$Start,

        [Parameter(Mandatory = $true)]
        [datetime]$End,

        [Parameter(Mandatory = $true)]
        [string]$ActivityType
    )

    $availableDates = @($DayMap.Values | ForEach-Object { $_['Date'].Date } | Sort-Object)
    $firstAvailableDate = $availableDates[0]
    $lastAvailableDate = $availableDates[-1]
    $cursor = $Start

    while ($cursor -lt $End) {
        $recordDate = if ($cursor.TimeOfDay -lt [TimeSpan]::FromHours(4)) {
            $cursor.Date.AddDays(-1)
        }
        else {
            $cursor.Date
        }

        $recordBoundary = $recordDate.AddDays(1).AddHours(4)
        $segmentEnd = if ($End -lt $recordBoundary) { $End } else { $recordBoundary }
        $dateKey = $recordDate.ToString('yyyy-MM-dd')

        if ($DayMap.ContainsKey($dateKey)) {
            $dayState = $DayMap[$dateKey]
            if ($null -eq $dayState -or $null -eq $dayState['Slots']) {
                throw "Slot buffer missing for $dateKey"
            }

            Add-DaySegmentToSlots -DayState $dayState -SegmentStart $cursor -SegmentEnd $segmentEnd -ActivityType $ActivityType
        }
        elseif ($recordDate -ge $firstAvailableDate -and $recordDate -le $lastAvailableDate) {
            throw "Activity '$ActivityType' covers $dateKey, but no matching day entry exists in JSON."
        }

        $cursor = $segmentEnd
    }
}

function Add-SleepDurationToDayMap {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$DayMap,

        [Parameter(Mandatory = $true)]
        [datetime]$Start,

        [Parameter(Mandatory = $true)]
        [datetime]$End
    )

    if ($End -le $Start) {
        return
    }

    foreach ($dayState in $DayMap.Values) {
        $windowStart = $dayState['Date'].Date.AddHours(16)
        $windowEnd = $dayState['Date'].Date.AddDays(1).AddHours(15)
        $overlapStart = if ($Start -gt $windowStart) { $Start } else { $windowStart }
        $overlapEnd = if ($End -lt $windowEnd) { $End } else { $windowEnd }

        if ($overlapEnd -gt $overlapStart) {
            $dayState['SleepMinutes'] = [int]$dayState['SleepMinutes'] + [int](($overlapEnd - $overlapStart).TotalMinutes)
        }
    }
}

function Get-ActivitySegments {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Slots
    )

    $segments = @()
    $slotIndex = 0
    while ($slotIndex -lt $Slots.Count) {
        $slotValue = $Slots[$slotIndex]
        if ([string]::IsNullOrWhiteSpace($slotValue)) {
            $slotIndex++
            continue
        }

        $startIndex = $slotIndex
        while (($slotIndex + 1) -lt $Slots.Count -and $Slots[$slotIndex + 1] -eq $slotValue) {
            $slotIndex++
        }

        $segments += [pscustomobject]@{
            Type      = $slotValue
            StartSlot = $startIndex
            EndSlot   = $slotIndex
        }

        $slotIndex++
    }

    return $segments
}

function Get-DatePages {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Days,

        [Parameter(Mandatory = $true)]
        [int]$PageSize
    )

    $pages = @()
    for ($offset = 0; $offset -lt $Days.Count; $offset += $PageSize) {
        $endIndex = [math]::Min($offset + $PageSize - 1, $Days.Count - 1)
        $pageDays = @($Days[$offset..$endIndex])
        $sheetName = '{0}_{1}' -f $pageDays[0]['Date'].ToString('yyyy-MM-dd'), $pageDays[-1]['Date'].ToString('yyyy-MM-dd')

        $pages += [pscustomobject]@{
            SheetName = $sheetName
            Days      = $pageDays
        }
    }

    return $pages
}

function Get-NextRelationshipId {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNodeList]$RelationshipNodes
    )

    $maxRelationshipNumber = 0
    foreach ($relationshipNode in $RelationshipNodes) {
        $match = [regex]::Match($relationshipNode.GetAttribute('Id'), '^rId(\d+)$')
        if ($match.Success) {
            $number = [int]$match.Groups[1].Value
            if ($number -gt $maxRelationshipNumber) {
                $maxRelationshipNumber = $number
            }
        }
    }

    return "rId$($maxRelationshipNumber + 1)"
}

function Quote-SheetNameForFormula {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SheetName
    )

    return "'{0}'" -f $SheetName.Replace("'", "''")
}

function Update-SheetRelationships {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$DrawingTarget
    )

    $document = Load-XmlDocument -Path $Path
    $drawingRelationship = @($document.DocumentElement.ChildNodes) |
        Where-Object { $_.LocalName -eq 'Relationship' -and $_.GetAttribute('Type').Contains('/drawing') } |
        Select-Object -First 1
    if (-not $drawingRelationship) {
        throw "Drawing relationship not found in $Path"
    }

    $drawingRelationship.SetAttribute('Target', $DrawingTarget)
    $document.Save($Path)
}

function Add-WorksheetOverride {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$ContentTypesDocument,

        [Parameter(Mandatory = $true)]
        [string]$PartName,

        [Parameter(Mandatory = $true)]
        [string]$ContentType
    )

    $existingNode = $ContentTypesDocument.SelectSingleNode("/ct:Types/ct:Override[@PartName='$PartName']", ([System.Xml.XmlNamespaceManager](New-NamespaceManager -Document $ContentTypesDocument)))
    if ($existingNode) {
        return
    }

    $overrideNode = $ContentTypesDocument.CreateElement('Override', $script:ContentTypesNamespace)
    $overrideNode.SetAttribute('PartName', $PartName)
    $overrideNode.SetAttribute('ContentType', $ContentType)
    [void]$ContentTypesDocument.DocumentElement.AppendChild($overrideNode)
}

function Update-WorkbookMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookRelationshipsPath,

        [Parameter(Mandatory = $true)]
        [string]$ContentTypesPath,

        [Parameter(Mandatory = $true)]
        [string[]]$SheetNames
    )

    $workbookDocument = Load-XmlDocument -Path $WorkbookPath
    $workbookNamespaceManager = [System.Xml.XmlNamespaceManager](New-NamespaceManager -Document $workbookDocument)
    $workbookRelationshipsDocument = Load-XmlDocument -Path $WorkbookRelationshipsPath
    $workbookRelationshipsNamespaceManager = [System.Xml.XmlNamespaceManager](New-NamespaceManager -Document $workbookRelationshipsDocument)
    $contentTypesDocument = Load-XmlDocument -Path $ContentTypesPath

    $sheetsNode = $workbookDocument.SelectSingleNode('/x:workbook/x:sheets', $workbookNamespaceManager)
    $sheetNodes = @($workbookDocument.SelectNodes('/x:workbook/x:sheets/x:sheet', $workbookNamespaceManager))
    if ($sheetNodes.Count -lt 1) {
        throw 'Workbook does not contain a worksheet node.'
    }

    $sheetNodes[0].SetAttribute('name', $SheetNames[0])

    $relationshipsRoot = $workbookRelationshipsDocument.SelectSingleNode('/pr:Relationships', $workbookRelationshipsNamespaceManager)
    $nextRelationshipId = Get-NextRelationshipId -RelationshipNodes $workbookRelationshipsDocument.SelectNodes('/pr:Relationships/pr:Relationship', $workbookRelationshipsNamespaceManager)
    $nextRelationshipNumber = [int]([regex]::Match($nextRelationshipId, '\d+').Value)

    for ($sheetIndex = 1; $sheetIndex -lt $SheetNames.Count; $sheetIndex++) {
        $relationshipId = "rId$nextRelationshipNumber"
        $nextRelationshipNumber++

        $sheetNode = $workbookDocument.CreateElement('sheet', $script:SpreadsheetNamespace)
        $sheetNode.SetAttribute('name', $SheetNames[$sheetIndex])
        $sheetNode.SetAttribute('sheetId', [string]($sheetIndex + 1))
        Set-NamespacedAttribute -Element $sheetNode -Prefix 'r' -LocalName 'id' -NamespaceUri $script:OfficeDocumentRelationshipNamespace -Value $relationshipId
        [void]$sheetsNode.AppendChild($sheetNode)

        $relationshipNode = $workbookRelationshipsDocument.CreateElement('Relationship', $script:PackageRelationshipNamespace)
        $relationshipNode.SetAttribute('Id', $relationshipId)
        $relationshipNode.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
        $relationshipNode.SetAttribute('Target', "worksheets/sheet$($sheetIndex + 1).xml")
        [void]$relationshipsRoot.AppendChild($relationshipNode)

        Add-WorksheetOverride -ContentTypesDocument $contentTypesDocument -PartName "/xl/worksheets/sheet$($sheetIndex + 1).xml" -ContentType 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
        Add-WorksheetOverride -ContentTypesDocument $contentTypesDocument -PartName "/xl/drawings/drawing$($sheetIndex + 1).xml" -ContentType 'application/vnd.openxmlformats-officedocument.drawing+xml'
    }

    $definedNamesNode = $workbookDocument.SelectSingleNode('/x:workbook/x:definedNames', $workbookNamespaceManager)
    if (-not $definedNamesNode) {
        $definedNamesNode = $workbookDocument.CreateElement('definedNames', $script:SpreadsheetNamespace)
        $calcPrNode = $workbookDocument.SelectSingleNode('/x:workbook/x:calcPr', $workbookNamespaceManager)
        if ($calcPrNode) {
            [void]$workbookDocument.DocumentElement.InsertBefore($definedNamesNode, $calcPrNode)
        }
        else {
            [void]$workbookDocument.DocumentElement.AppendChild($definedNamesNode)
        }
    }

    foreach ($printAreaNode in @($workbookDocument.SelectNodes("/x:workbook/x:definedNames/x:definedName[@name='_xlnm.Print_Area']", $workbookNamespaceManager))) {
        [void]$definedNamesNode.RemoveChild($printAreaNode)
    }

    for ($sheetIndex = 0; $sheetIndex -lt $SheetNames.Count; $sheetIndex++) {
        $definedNameNode = $workbookDocument.CreateElement('definedName', $script:SpreadsheetNamespace)
        $definedNameNode.SetAttribute('name', '_xlnm.Print_Area')
        $definedNameNode.SetAttribute('localSheetId', [string]$sheetIndex)
        $definedNameNode.InnerText = "{0}!{1}" -f (Quote-SheetNameForFormula -SheetName $SheetNames[$sheetIndex]), $script:PrintableArea
        [void]$definedNamesNode.AppendChild($definedNameNode)
    }

    $workbookDocument.Save($WorkbookPath)
    $workbookRelationshipsDocument.Save($WorkbookRelationshipsPath)
    $contentTypesDocument.Save($ContentTypesPath)
}

function Update-AppProperties {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppPropertiesPath,

        [Parameter(Mandatory = $true)]
        [string[]]$SheetNames
    )

    $appDocument = Load-XmlDocument -Path $AppPropertiesPath
    $namespaceManager = [System.Xml.XmlNamespaceManager](New-NamespaceManager -Document $appDocument)

    $headingCounts = @($appDocument.SelectNodes('/ap:Properties/ap:HeadingPairs/vt:vector/vt:variant/vt:i4', $namespaceManager))
    if ($headingCounts.Count -ge 2) {
        $headingCounts[0].InnerText = [string]$SheetNames.Count
        $headingCounts[1].InnerText = [string]$SheetNames.Count
    }

    $titlesVectorNode = $appDocument.SelectSingleNode('/ap:Properties/ap:TitlesOfParts/vt:vector', $namespaceManager)
    if (-not $titlesVectorNode) {
        throw "TitlesOfParts vector not found in $AppPropertiesPath"
    }

    while ($titlesVectorNode.HasChildNodes) {
        [void]$titlesVectorNode.RemoveChild($titlesVectorNode.FirstChild)
    }

    foreach ($sheetName in $SheetNames) {
        $sheetTitleNode = $appDocument.CreateElement('vt', 'lpstr', $script:VariantTypesNamespace)
        $sheetTitleNode.InnerText = $sheetName
        [void]$titlesVectorNode.AppendChild($sheetTitleNode)
    }

    foreach ($sheetName in $SheetNames) {
        $printAreaTitleNode = $appDocument.CreateElement('vt', 'lpstr', $script:VariantTypesNamespace)
        $printAreaTitleNode.InnerText = "$sheetName!Print_Area"
        [void]$titlesVectorNode.AppendChild($printAreaTitleNode)
    }

    $titlesVectorNode.SetAttribute('size', [string]($SheetNames.Count * 2))
    $appDocument.Save($AppPropertiesPath)
}

function Update-CoreProperties {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CorePropertiesPath
    )

    $coreDocument = Load-XmlDocument -Path $CorePropertiesPath
    $namespaceManager = [System.Xml.XmlNamespaceManager](New-NamespaceManager -Document $coreDocument)
    $modifiedUtc = [DateTime]::UtcNow.ToString('s') + 'Z'

    $lastModifiedByNode = $coreDocument.SelectSingleNode('/cp:coreProperties/cp:lastModifiedBy', $namespaceManager)
    if ($lastModifiedByNode) {
        $lastModifiedByNode.InnerText = 'Codex'
    }

    $modifiedNode = $coreDocument.SelectSingleNode('/cp:coreProperties/dcterms:modified', $namespaceManager)
    if ($modifiedNode) {
        $modifiedNode.InnerText = $modifiedUtc
    }

    $coreDocument.Save($CorePropertiesPath)
}

function Add-BlackFillStyles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StylesPath
    )

    $stylesDocument = Load-XmlDocument -Path $StylesPath
    $namespaceManager = [System.Xml.XmlNamespaceManager](New-NamespaceManager -Document $stylesDocument)
    $fillsNode = $stylesDocument.SelectSingleNode('/x:styleSheet/x:fills', $namespaceManager)
    $cellXfsNode = $stylesDocument.SelectSingleNode('/x:styleSheet/x:cellXfs', $namespaceManager)
    if (-not $fillsNode -or -not $cellXfsNode) {
        throw "Styles workbook part is missing fills or cellXfs: $StylesPath"
    }

    $blackFillId = @($fillsNode.SelectNodes('./x:fill', $namespaceManager)).Count
    $fillNode = $stylesDocument.CreateElement('fill', $script:SpreadsheetNamespace)
    $patternFillNode = $stylesDocument.CreateElement('patternFill', $script:SpreadsheetNamespace)
    $patternFillNode.SetAttribute('patternType', 'solid')
    $foregroundColorNode = $stylesDocument.CreateElement('fgColor', $script:SpreadsheetNamespace)
    $foregroundColorNode.SetAttribute('rgb', 'FF000000')
    $backgroundColorNode = $stylesDocument.CreateElement('bgColor', $script:SpreadsheetNamespace)
    $backgroundColorNode.SetAttribute('indexed', '64')
    [void]$patternFillNode.AppendChild($foregroundColorNode)
    [void]$patternFillNode.AppendChild($backgroundColorNode)
    [void]$fillNode.AppendChild($patternFillNode)
    [void]$fillsNode.AppendChild($fillNode)
    $fillsNode.SetAttribute('count', [string]$fillsNode.ChildNodes.Count)

    $blackFillStyleIndexes = @{}
    $originalStyleNodes = @($cellXfsNode.SelectNodes('./x:xf', $namespaceManager))
    for ($styleIndex = 0; $styleIndex -lt $originalStyleNodes.Count; $styleIndex++) {
        $blackStyleIndex = $cellXfsNode.ChildNodes.Count
        $blackStyleNode = [System.Xml.XmlElement]$originalStyleNodes[$styleIndex].CloneNode($true)
        $blackStyleNode.SetAttribute('fillId', [string]$blackFillId)
        $blackStyleNode.SetAttribute('applyFill', '1')
        [void]$cellXfsNode.AppendChild($blackStyleNode)
        $blackFillStyleIndexes[[string]$styleIndex] = [string]$blackStyleIndex
    }

    $cellXfsNode.SetAttribute('count', [string]$cellXfsNode.ChildNodes.Count)
    $stylesDocument.Save($StylesPath)

    return $blackFillStyleIndexes
}

function Update-WorksheetData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorksheetPath,

        [Parameter(Mandatory = $true)]
        [string]$SheetName,

        [Parameter(Mandatory = $true)]
        [object[]]$Days,

        [Parameter(Mandatory = $true)]
        [hashtable]$BlackFillStyleIndexes,

        [string]$NameNo,

        [string]$Name
    )

    $worksheetDocument = Load-XmlDocument -Path $WorksheetPath
    $namespaceManager = [System.Xml.XmlNamespaceManager](New-NamespaceManager -Document $worksheetDocument)

    if ($Days.Count -lt 1) {
        throw "Worksheet $SheetName does not contain any day data."
    }

    $headerRowNode = $worksheetDocument.SelectSingleNode('/x:worksheet/x:sheetData/x:row[@r=''1'']', $namespaceManager)
    if (-not $headerRowNode) {
        throw "Row 1 not found in worksheet $SheetName"
    }

    $oldestDate = [datetime]$Days[0]['Date']
    $newestDate = [datetime]$Days[-1]['Date']
    Set-CellNumber -RowNode $headerRowNode -NamespaceManager $namespaceManager -CellReference 'K1' -Value $oldestDate.Year
    Set-CellNumber -RowNode $headerRowNode -NamespaceManager $namespaceManager -CellReference 'O1' -Value $oldestDate.Month
    Set-CellNumber -RowNode $headerRowNode -NamespaceManager $namespaceManager -CellReference 'R1' -Value $oldestDate.Day
    Set-CellNumber -RowNode $headerRowNode -NamespaceManager $namespaceManager -CellReference 'W1' -Value $newestDate.Month
    Set-CellNumber -RowNode $headerRowNode -NamespaceManager $namespaceManager -CellReference 'Z1' -Value $newestDate.Day

    if (-not [string]::IsNullOrEmpty($NameNo)) {
        Set-CellInlineString -RowNode $headerRowNode -NamespaceManager $namespaceManager -CellReference 'AK1' -Value $NameNo
    }

    if (-not [string]::IsNullOrEmpty($Name)) {
        Set-CellInlineString -RowNode $headerRowNode -NamespaceManager $namespaceManager -CellReference 'AX1' -Value $Name
    }

    $sheetViewNode = $worksheetDocument.SelectSingleNode('/x:worksheet/x:sheetViews/x:sheetView', $namespaceManager)
    if ($sheetViewNode) {
        $sheetViewNode.RemoveAttribute('tabSelected')
    }

    Set-NamespacedAttribute -Element $worksheetDocument.DocumentElement -Prefix 'xr' -LocalName 'uid' -NamespaceUri $script:SpreadsheetRevisionNamespace -Value ("{" + ([guid]::NewGuid().ToString().ToUpperInvariant()) + "}")

    $dimensionNode = $worksheetDocument.SelectSingleNode('/x:worksheet/x:dimension', $namespaceManager)
    if ($dimensionNode) {
        $dimensionNode.SetAttribute('ref', $script:DimensionRef)
    }

    $mergeCellsNode = $worksheetDocument.SelectSingleNode('/x:worksheet/x:mergeCells', $namespaceManager)
    if (-not $mergeCellsNode) {
        throw "mergeCells node not found in $WorksheetPath"
    }

    foreach ($dayIndex in 0..($Days.Count - 1)) {
        $day = $Days[$dayIndex]
        $rowNumber = $script:DayStartRows[$dayIndex]
        $rowNode = $worksheetDocument.SelectSingleNode("/x:worksheet/x:sheetData/x:row[@r='$rowNumber']", $namespaceManager)
        if (-not $rowNode) {
            throw "Row $rowNumber not found in worksheet $SheetName"
        }

        Set-CellNumber -RowNode $rowNode -NamespaceManager $namespaceManager -CellReference "A$rowNumber" -Value $day['Date'].Month -StyleIndex '18'
        Set-CellNumber -RowNode $rowNode -NamespaceManager $namespaceManager -CellReference "C$rowNumber" -Value $day['Date'].Day -StyleIndex '18'
        Set-CellInlineString -RowNode $rowNode -NamespaceManager $namespaceManager -CellReference "F$rowNumber" -Value (Get-DayOfWeekLabel -Date $day['Date']) -StyleIndex '13'

        if (-not [string]::IsNullOrWhiteSpace($day['Note'])) {
            Set-CellInlineString -RowNode $rowNode -NamespaceManager $namespaceManager -CellReference "BH$rowNumber" -Value $day['Note']
            Set-RowMaxSpan -RowNode $rowNode -MaxColumnNumber $script:NoteColumnNumber
        }

        foreach ($segment in Get-ActivitySegments -Slots $day['Slots']) {
            $segmentStartColumnNumber = $script:TimeSlotStartColumn + $segment.StartSlot
            $segmentEndColumnNumber = $script:TimeSlotStartColumn + $segment.EndSlot
            $isSleepSegment = $segment.Type -eq $script:SleepActivityType

            if ($isSleepSegment) {
                for ($columnNumber = $segmentStartColumnNumber; $columnNumber -le $segmentEndColumnNumber; $columnNumber++) {
                    $sleepCellReference = "{0}{1}" -f (ConvertTo-ColumnName -ColumnNumber $columnNumber), $rowNumber
                    $sleepCellNode = Get-OrCreate-CellNode -RowNode $rowNode -NamespaceManager $namespaceManager -CellReference $sleepCellReference
                    $baseStyleIndex = $sleepCellNode.GetAttribute('s')
                    if ([string]::IsNullOrWhiteSpace($baseStyleIndex)) {
                        $baseStyleIndex = '0'
                    }

                    if ($BlackFillStyleIndexes.ContainsKey($baseStyleIndex)) {
                        Set-CellStyle -RowNode $rowNode -NamespaceManager $namespaceManager -CellReference $sleepCellReference -StyleIndex $BlackFillStyleIndexes[$baseStyleIndex]
                    }
                }
            }

            $startColumn = ConvertTo-ColumnName -ColumnNumber $segmentStartColumnNumber
            $endColumn = ConvertTo-ColumnName -ColumnNumber $segmentEndColumnNumber
            $startCellReference = "$startColumn$rowNumber"
            Set-CellInlineString -RowNode $rowNode -NamespaceManager $namespaceManager -CellReference $startCellReference -Value $segment.Type

            if ($segmentStartColumnNumber -lt $segmentEndColumnNumber) {
                $mergeCellNode = $worksheetDocument.CreateElement('mergeCell', $script:SpreadsheetNamespace)
                $mergeCellNode.SetAttribute('ref', ('{0}{1}:{2}{1}' -f $startColumn, $rowNumber, $endColumn))
                [void]$mergeCellsNode.AppendChild($mergeCellNode)
            }
        }

        $sleepHours = [double]$day['SleepMinutes'] / 60
        Set-CellDecimalNumber -RowNode $rowNode -NamespaceManager $namespaceManager -CellReference "BF$rowNumber" -Value $sleepHours
        Set-RowMaxSpan -RowNode $rowNode -MaxColumnNumber $script:SleepTotalColumnNumber

        $pcOffRowNumber = $rowNumber - 1
        $pcOffRowNode = $worksheetDocument.SelectSingleNode("/x:worksheet/x:sheetData/x:row[@r='$pcOffRowNumber']", $namespaceManager)
        if (-not $pcOffRowNode) {
            throw "Row $pcOffRowNumber not found in worksheet $SheetName"
        }

        Set-CellInlineString -RowNode $pcOffRowNode -NamespaceManager $namespaceManager -CellReference "AR$pcOffRowNumber" -Value $script:PcOffText
        Set-RowMaxSpan -RowNode $pcOffRowNode -MaxColumnNumber $script:PcOffColumnNumber
    }

    $mergeCellsNode.SetAttribute('count', [string]$mergeCellsNode.ChildNodes.Count)
    $worksheetDocument.Save($WorksheetPath)
}

$inputJsonPath = Resolve-AbsolutePath -Path $InputJson -MustExist
$templateXlsxPath = Resolve-AbsolutePath -Path $TemplateXlsx -MustExist
if (-not $OutputXlsx) {
    $OutputXlsx = Join-Path -Path (Split-Path -Parent $templateXlsxPath) -ChildPath 'seikatsu_converted.xlsx'
}

$outputXlsxPath = Resolve-AbsolutePath -Path $OutputXlsx

if ($inputJsonPath -eq $templateXlsxPath) {
    throw 'Input JSON path and template xlsx path must be different files.'
}

if ($templateXlsxPath -eq $outputXlsxPath) {
    throw 'Output xlsx path must be different from the template xlsx path.'
}

$outputDirectory = Split-Path -Parent $outputXlsxPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
}

$jsonText = Get-Content -LiteralPath $inputJsonPath -Encoding UTF8 -Raw
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = 100000000
$jsonData = $serializer.DeserializeObject($jsonText)

if (-not $jsonData.ContainsKey('days') -or -not $jsonData.ContainsKey('activities')) {
    throw 'The source JSON must contain both days and activities arrays.'
}

$invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
$days = @($jsonData['days'] | Sort-Object { [datetime]::ParseExact($_['date'], 'yyyy-MM-dd', $invariantCulture) })
$pageSize = $script:DayStartRows.Count

$dayMap = @{}
foreach ($dayEntry in $days) {
    $date = [datetime]::ParseExact($dayEntry['date'], 'yyyy-MM-dd', $invariantCulture).Date
    $dayKey = $date.ToString('yyyy-MM-dd')
    $slots = [System.Collections.Generic.List[string]]::new()
    foreach ($slotIndex in 1..$script:HalfHourSlotsPerDay) {
        [void]$slots.Add('')
    }

    $dayMap[$dayKey] = @{
        Date         = $date
        Note         = if ($dayEntry.ContainsKey('note')) { [string]$dayEntry['note'] } else { '' }
        Slots        = $slots
        SleepMinutes = 0
    }
}

foreach ($activityEntry in @($jsonData['activities'] | Sort-Object { [datetime]::Parse($_['start'], $invariantCulture) })) {
    $activityType = [string]$activityEntry['type']
    if ([string]::IsNullOrWhiteSpace($activityType)) {
        continue
    }

    $activityStart = [datetime]::Parse($activityEntry['start'], $invariantCulture)
    $activityEnd = $activityStart.AddMinutes([int]$activityEntry['durationMinutes'])

    $roundedStart = Floor-ToHalfHour -DateTime $activityStart
    $roundedEnd = Ceiling-ToHalfHour -DateTime $activityEnd
    if ($roundedEnd -le $roundedStart) {
        continue
    }

    Add-ActivityToDayMap -DayMap $dayMap -Start $roundedStart -End $roundedEnd -ActivityType $activityType
    if ($activityType -eq $script:SleepActivityType) {
        Add-SleepDurationToDayMap -DayMap $dayMap -Start $roundedStart -End $roundedEnd
    }
}

$orderedDayObjects = foreach ($dayEntry in $days) {
    $dayMap[$dayEntry['date']]
}

$pages = Get-DatePages -Days $orderedDayObjects -PageSize $pageSize
$sheetNames = @($pages | ForEach-Object { $_.SheetName })

$workspaceTempRoot = Join-Path -Path $outputDirectory -ChildPath ('.tmp-activity-log-' + [guid]::NewGuid().ToString('N'))
$expandedWorkbookDirectory = Join-Path -Path $workspaceTempRoot -ChildPath 'expanded'
$temporaryOutputPath = Join-Path -Path $workspaceTempRoot -ChildPath 'result.xlsx'

try {
    [void](New-Item -ItemType Directory -Path $expandedWorkbookDirectory -Force)
    [System.IO.Compression.ZipFile]::ExtractToDirectory($templateXlsxPath, $expandedWorkbookDirectory)

    $worksheetsDirectory = Join-Path -Path $expandedWorkbookDirectory -ChildPath 'xl\worksheets'
    $worksheetRelationshipsDirectory = Join-Path -Path $worksheetsDirectory -ChildPath '_rels'
    $drawingsDirectory = Join-Path -Path $expandedWorkbookDirectory -ChildPath 'xl\drawings'

    if ($sheetNames.Count -gt 1) {
        foreach ($sheetNumber in 2..$sheetNames.Count) {
            Copy-Item -LiteralPath (Join-Path $worksheetsDirectory 'sheet1.xml') -Destination (Join-Path $worksheetsDirectory "sheet$sheetNumber.xml")
            Copy-Item -LiteralPath (Join-Path $worksheetRelationshipsDirectory 'sheet1.xml.rels') -Destination (Join-Path $worksheetRelationshipsDirectory "sheet$sheetNumber.xml.rels")
            Copy-Item -LiteralPath (Join-Path $drawingsDirectory 'drawing1.xml') -Destination (Join-Path $drawingsDirectory "drawing$sheetNumber.xml")

            Update-SheetRelationships -Path (Join-Path $worksheetRelationshipsDirectory "sheet$sheetNumber.xml.rels") -DrawingTarget "../drawings/drawing$sheetNumber.xml"
        }
    }

    Update-WorkbookMetadata `
        -WorkbookPath (Join-Path $expandedWorkbookDirectory 'xl\workbook.xml') `
        -WorkbookRelationshipsPath (Join-Path $expandedWorkbookDirectory 'xl\_rels\workbook.xml.rels') `
        -ContentTypesPath (Join-Path $expandedWorkbookDirectory '[Content_Types].xml') `
        -SheetNames $sheetNames

    Update-AppProperties -AppPropertiesPath (Join-Path $expandedWorkbookDirectory 'docProps\app.xml') -SheetNames $sheetNames
    Update-CoreProperties -CorePropertiesPath (Join-Path $expandedWorkbookDirectory 'docProps\core.xml')
    $blackFillStyleIndexes = Add-BlackFillStyles -StylesPath (Join-Path $expandedWorkbookDirectory 'xl\styles.xml')

    for ($pageIndex = 0; $pageIndex -lt $pages.Count; $pageIndex++) {
        $worksheetPath = Join-Path -Path $worksheetsDirectory -ChildPath "sheet$($pageIndex + 1).xml"
        Update-WorksheetData -WorksheetPath $worksheetPath -SheetName $pages[$pageIndex].SheetName -Days $pages[$pageIndex].Days -BlackFillStyleIndexes $blackFillStyleIndexes -NameNo $NameNo -Name $Name
    }

    Write-ZipArchiveFromDirectory -SourceDirectory $expandedWorkbookDirectory -DestinationFile $temporaryOutputPath
    Copy-Item -LiteralPath $temporaryOutputPath -Destination $outputXlsxPath -Force
}
finally {
    if (Test-Path -LiteralPath $workspaceTempRoot) {
        $removed = $false
        foreach ($attempt in 1..3) {
            try {
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
                Start-Sleep -Milliseconds 200
                Remove-Item -LiteralPath $workspaceTempRoot -Recurse -Force
                $removed = $true
                break
            }
            catch {
                continue
            }
        }

        if (-not $removed -and (Test-Path -LiteralPath $workspaceTempRoot)) {
            Write-Verbose ("Temporary workspace kept for inspection: {0}" -f $workspaceTempRoot)
        }
    }
}

Write-Output ("Converted {0} day(s) into {1} sheet(s): {2}" -f $orderedDayObjects.Count, $pages.Count, $outputXlsxPath)

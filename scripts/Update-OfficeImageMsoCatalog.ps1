[CmdletBinding()]
param(
    [string] $OutputPath,

    [string] $ChannelPath = 'Microsoft 365/Current Channel'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$RepositoryApiRoot = 'https://api.github.com/repos/OfficeDev/office-fluent-ui-command-identifiers'

function Get-ColumnIndex {
    param(
        [Parameter(Mandatory = $true)]
        [string] $CellReference
    )

    $letters = ([regex]::Match($CellReference, '^[A-Z]+')).Value
    $index = 0
    foreach ($character in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$character - [int][char]'A' + 1)
    }

    return $index
}

function Get-EntryText {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive] $Archive,

        [Parameter(Mandatory = $true)]
        [string] $EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        return $null
    }

    $stream = $null
    $reader = $null
    try {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        return $reader.ReadToEnd()
    } finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        } elseif ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Get-SharedStrings {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive] $Archive
    )

    $content = Get-EntryText -Archive $Archive -EntryName 'xl/sharedStrings.xml'
    if ([string]::IsNullOrWhiteSpace($content)) {
        return @()
    }

    $document = New-Object System.Xml.XmlDocument
    $document.LoadXml($content)

    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
    $namespaceManager.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $strings = New-Object System.Collections.Generic.List[string]
    foreach ($item in $document.SelectNodes('/x:sst/x:si', $namespaceManager)) {
        $textParts = foreach ($textNode in $item.SelectNodes('.//x:t', $namespaceManager)) {
            $textNode.InnerText
        }

        $strings.Add(($textParts -join ''))
    }

    return $strings.ToArray()
}

function Get-CellValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement] $Cell,

        [Parameter(Mandatory = $true)]
        [string[]] $SharedStrings,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNamespaceManager] $NamespaceManager
    )

    $type = $Cell.GetAttribute('t')
    if ($type -eq 'inlineStr') {
        return (($Cell.SelectNodes('.//x:t', $NamespaceManager) | ForEach-Object { $_.InnerText }) -join '')
    }

    $valueNode = $Cell.SelectSingleNode('x:v', $NamespaceManager)
    if ($null -eq $valueNode) {
        return ''
    }

    if ($type -eq 's') {
        $index = [int] $valueNode.InnerText
        if ($index -ge 0 -and $index -lt $SharedStrings.Count) {
            return $SharedStrings[$index]
        }

        return ''
    }

    return $valueNode.InnerText
}

function Read-FirstWorksheet {
    param(
        [Parameter(Mandatory = $true)]
        [string] $WorkbookPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
        $sharedStrings = Get-SharedStrings -Archive $archive
        $worksheetEntry = $archive.Entries |
            Where-Object { $_.FullName -like 'xl/worksheets/*.xml' } |
            Sort-Object FullName |
            Select-Object -First 1

        if ($null -eq $worksheetEntry) {
            throw "Workbook contains no worksheets: $WorkbookPath"
        }

        $worksheetXml = Get-EntryText -Archive $archive -EntryName $worksheetEntry.FullName
        $document = New-Object System.Xml.XmlDocument
        $document.LoadXml($worksheetXml)

        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
        $namespaceManager.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

        $rows = New-Object System.Collections.Generic.List[object[]]
        foreach ($row in $document.SelectNodes('/x:worksheet/x:sheetData/x:row', $namespaceManager)) {
            $valuesByColumn = @{}
            $maxColumn = 0
            foreach ($cell in $row.SelectNodes('x:c', $namespaceManager)) {
                $columnIndex = Get-ColumnIndex -CellReference $cell.GetAttribute('r')
                $valuesByColumn[$columnIndex] = Get-CellValue -Cell $cell -SharedStrings $sharedStrings -NamespaceManager $namespaceManager
                if ($columnIndex -gt $maxColumn) {
                    $maxColumn = $columnIndex
                }
            }

            $values = for ($i = 1; $i -le $maxColumn; $i++) {
                if ($valuesByColumn.ContainsKey($i)) {
                    $valuesByColumn[$i]
                } else {
                    ''
                }
            }

            $rows.Add([object[]] $values)
        }

        return $rows.ToArray()
    } finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }
    }
}

function Get-Value {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $Row,

        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    if ($Row.ContainsKey($Name)) {
        return ($Row[$Name]).Trim()
    }

    return ''
}

function Get-AppName {
    param(
        [Parameter(Mandatory = $true)]
        [string] $FileName
    )

    $name = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    return ($name -replace 'controls$', '')
}

function Join-Description {
    param(
        [string] $Application,
        [string] $ControlType,
        [string] $Tab,
        [string] $GroupOrContextMenu,
        [string] $ParentControl,
        [string] $SecondaryParentControl
    )

    $parts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($Application)) {
        $parts.Add("App: $Application")
    }

    if (-not [string]::IsNullOrWhiteSpace($ControlType)) {
        $parts.Add("Typ: $ControlType")
    }

    if (-not [string]::IsNullOrWhiteSpace($Tab)) {
        $parts.Add("Tab: $Tab")
    }

    if (-not [string]::IsNullOrWhiteSpace($GroupOrContextMenu)) {
        $parts.Add("Gruppe/Kontext: $GroupOrContextMenu")
    }

    if (-not [string]::IsNullOrWhiteSpace($ParentControl)) {
        $parts.Add("Parent: $ParentControl")
    }

    if (-not [string]::IsNullOrWhiteSpace($SecondaryParentControl)) {
        $parts.Add("Secondary Parent: $SecondaryParentControl")
    }

    return ($parts -join '; ')
}

function Convert-WorkbookToCatalogRows {
    param(
        [Parameter(Mandatory = $true)]
        [string] $WorkbookPath,

        [Parameter(Mandatory = $true)]
        [string] $SourceFile
    )

    $rows = Read-FirstWorksheet -WorkbookPath $WorkbookPath
    if ($rows.Count -lt 2) {
        return @()
    }

    $headers = @($rows[0])
    $application = Get-AppName -FileName $SourceFile
    $catalogRows = New-Object System.Collections.Generic.List[object]

    for ($rowIndex = 1; $rowIndex -lt $rows.Count; $rowIndex++) {
        $values = @($rows[$rowIndex])
        $row = @{}
        for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
            if ($columnIndex -lt $values.Count) {
                $row[$headers[$columnIndex]] = [string] $values[$columnIndex]
            } else {
                $row[$headers[$columnIndex]] = ''
            }
        }

        $controlName = Get-Value -Row $row -Name 'Control Name'
        if ([string]::IsNullOrWhiteSpace($controlName)) {
            continue
        }

        $controlType = Get-Value -Row $row -Name 'Control Type'
        $tabSet = Get-Value -Row $row -Name 'Tab Set'
        $tab = Get-Value -Row $row -Name 'Tab'
        $groupOrContextMenu = Get-Value -Row $row -Name 'Group/Context Menu Name'
        $parentControl = Get-Value -Row $row -Name 'Parent Control'
        $secondaryParentControl = Get-Value -Row $row -Name 'Secondary Parent Control'
        $policyId = Get-Value -Row $row -Name 'Policy ID'

        $catalogRows.Add([pscustomobject] @{
            ImageMso = $controlName
            Application = $application
            SourceFile = $SourceFile
            ControlType = $controlType
            TabSet = $tabSet
            Tab = $tab
            GroupOrContextMenu = $groupOrContextMenu
            ParentControl = $parentControl
            SecondaryParentControl = $secondaryParentControl
            PolicyId = $policyId
            Description = Join-Description `
                -Application $application `
                -ControlType $controlType `
                -Tab $tab `
                -GroupOrContextMenu $groupOrContextMenu `
                -ParentControl $parentControl `
                -SecondaryParentControl $secondaryParentControl
        })
    }

    return $catalogRows.ToArray()
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $PSScriptRoot '..\docs\OfficeImageMso-Catalog.csv'
}

$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outputDirectory = [System.IO.Path]::GetDirectoryName($resolvedOutputPath)
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

$encodedChannelPath = [System.Uri]::EscapeDataString($ChannelPath).Replace('%2F', '/')
$contentsUrl = "$RepositoryApiRoot/contents/$encodedChannelPath`?ref=main"
$items = Invoke-RestMethod -Uri $contentsUrl -Headers @{ 'User-Agent' = 'OfficeRibbonXEditor-CatalogUpdater' }
$workbookItems = @($items | Where-Object { $_.type -eq 'file' -and $_.name -like '*controls.xlsx' } | Sort-Object name)

if ($workbookItems.Count -eq 0) {
    throw "No *controls.xlsx files found at $ChannelPath."
}

$downloadDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ("OfficeImageMsoCatalog-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $downloadDirectory | Out-Null

$catalogRows = New-Object System.Collections.Generic.List[object]
try {
    foreach ($item in $workbookItems) {
        $localWorkbookPath = Join-Path $downloadDirectory $item.name
        Write-Verbose "Downloading $($item.download_url)"
        Invoke-WebRequest -Uri $item.download_url -OutFile $localWorkbookPath

        foreach ($catalogRow in (Convert-WorkbookToCatalogRows -WorkbookPath $localWorkbookPath -SourceFile $item.name)) {
            $catalogRows.Add($catalogRow)
        }
    }
} finally {
    if (Test-Path -LiteralPath $downloadDirectory -PathType Container) {
        Remove-Item -LiteralPath $downloadDirectory -Recurse -Force
    }
}

$catalogRows |
    Sort-Object ImageMso, Application, Tab, GroupOrContextMenu |
    ConvertTo-Csv -NoTypeInformation |
    Set-Content -LiteralPath $resolvedOutputPath -Encoding UTF8

Write-Host "Office imageMso catalog written to $resolvedOutputPath"
Write-Host "Rows: $($catalogRows.Count)"
Write-Host "Source: https://github.com/OfficeDev/office-fluent-ui-command-identifiers/tree/main/$($ChannelPath.Replace(' ', '%20'))"

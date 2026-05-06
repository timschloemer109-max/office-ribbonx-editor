[CmdletBinding()]
param(
    [string] $OfficePath,
    [string] $CustomUiPath,
    [string] $OutputPath,
    [switch] $ForceInPlace,
    [switch] $Interactive
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$CustomUi14Namespace = 'http://schemas.microsoft.com/office/2009/07/customui'
$CustomUiRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
$PackageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'

function Assert-SupportedOfficePath {
    param([string] $Path, [string] $Name)
    $supportedExtensions = @('.xlsm', '.xlam', '.docm', '.dotm', '.pptm', '.ppam')
    $extension = [System.IO.Path]::GetExtension($Path)
    if ($supportedExtensions -notcontains $extension.ToLowerInvariant()) {
        throw "$Name muss eine unterstuetzte Office-Makrodatei sein: $Path"
    }
}

function Resolve-RequiredFile {
    param([string] $Path, [string] $Name)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Name wurde nicht gefunden: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).ProviderPath
}

function Select-OfficeFile {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Office-Datei auswaehlen, die aktualisiert werden soll'
    $dialog.Filter = 'Office Makrodateien (*.xlsm;*.xlam;*.docm;*.dotm;*.pptm;*.ppam)|*.xlsm;*.xlam;*.docm;*.dotm;*.pptm;*.ppam|Alle Dateien (*.*)|*.*'
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw 'Keine Office-Datei ausgewaehlt.'
    }
    return $dialog.FileName
}

function Select-CustomUiFile {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'customUI14.txt oder customUI14.xml auswaehlen'
    $dialog.Filter = 'CustomUI14 XML (*.txt;*.xml)|*.txt;*.xml|Alle Dateien (*.*)|*.*'
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw 'Keine CustomUI-Datei ausgewaehlt.'
    }
    return $dialog.FileName
}

function Select-OutputFile {
    param([string] $SourcePath)
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = 'Ausgabedatei speichern'
    $dialog.Filter = 'Office Makrodateien (*.xlsm;*.xlam;*.docm;*.dotm;*.pptm;*.ppam)|*.xlsm;*.xlam;*.docm;*.dotm;*.pptm;*.ppam|Alle Dateien (*.*)|*.*'
    $dialog.FileName = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath) + '.customui' + [System.IO.Path]::GetExtension($SourcePath)
    $dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($SourcePath)
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw 'Keine Ausgabedatei ausgewaehlt.'
    }
    return $dialog.FileName
}

function Read-ZipText {
    param([System.IO.Compression.ZipArchive] $Zip, [string] $EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-ZipText {
    param([System.IO.Compression.ZipArchive] $Zip, [string] $EntryName, [string] $Text)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -ne $entry) { $entry.Delete() }
    $newEntry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter($newEntry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Assert-CustomUiXml {
    param([string] $XmlText)
    $document = New-Object System.Xml.XmlDocument
    $document.PreserveWhitespace = $true
    $document.LoadXml($XmlText)
    if ($document.DocumentElement.LocalName -ne 'customUI' -or $document.DocumentElement.NamespaceURI -ne $CustomUi14Namespace) {
        throw "CustomUI-Datei muss ein customUI14-XML mit Namespace $CustomUi14Namespace sein."
    }
}

function Ensure-XmlContentType {
    param([System.IO.Compression.ZipArchive] $Zip)
    [xml] $document = Read-ZipText -Zip $Zip -EntryName '[Content_Types].xml'
    $namespace = 'http://schemas.openxmlformats.org/package/2006/content-types'
    $hasXmlDefault = $false
    foreach ($default in $document.Types.Default) {
        if ($default.Extension -eq 'xml') { $hasXmlDefault = $true }
    }
    if (-not $hasXmlDefault) {
        $node = $document.CreateElement('Default', $namespace)
        $node.SetAttribute('Extension', 'xml')
        $node.SetAttribute('ContentType', 'application/xml')
        [void] $document.Types.AppendChild($node)
        Write-ZipText -Zip $Zip -EntryName '[Content_Types].xml' -Text $document.OuterXml
    }
}

function Ensure-CustomUiRelationship {
    param([System.IO.Compression.ZipArchive] $Zip)
    $relationshipEntry = '_rels/.rels'
    $text = Read-ZipText -Zip $Zip -EntryName $relationshipEntry
    if ([string]::IsNullOrWhiteSpace($text)) {
        $text = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + $PackageRelationshipNamespace + '"></Relationships>'
    }
    [xml] $document = $text
    $hasRelationship = $false
    $maxId = 0
    foreach ($relationship in $document.Relationships.Relationship) {
        if ($relationship.Type -eq $CustomUiRelationshipType -and $relationship.Target.TrimStart('/') -ieq 'customUI/customUI14.xml') {
            $hasRelationship = $true
        }
        if ($relationship.Id -match '^rId(\d+)$') {
            $maxId = [Math]::Max($maxId, [int] $Matches[1])
        }
    }
    if (-not $hasRelationship) {
        $node = $document.CreateElement('Relationship', $PackageRelationshipNamespace)
        $node.SetAttribute('Id', 'rId' + ($maxId + 1))
        $node.SetAttribute('Type', $CustomUiRelationshipType)
        $node.SetAttribute('Target', 'customUI/customUI14.xml')
        [void] $document.Relationships.AppendChild($node)
        Write-ZipText -Zip $Zip -EntryName $relationshipEntry -Text $document.OuterXml
    }
}

function Set-CustomUi14Xml {
    param([string] $TargetPath, [string] $XmlText)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($TargetPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        Ensure-XmlContentType -Zip $zip
        Ensure-CustomUiRelationship -Zip $zip
        Write-ZipText -Zip $zip -EntryName 'customUI/customUI14.xml' -Text $XmlText
    }
    finally {
        $zip.Dispose()
    }
}

try {
    if ($Interactive) {
        if ([string]::IsNullOrWhiteSpace($OfficePath)) { $OfficePath = Select-OfficeFile }
        if ([string]::IsNullOrWhiteSpace($CustomUiPath)) { $CustomUiPath = Select-CustomUiFile }
        if (-not $ForceInPlace -and [string]::IsNullOrWhiteSpace($OutputPath)) {
            $OutputPath = Select-OutputFile -SourcePath $OfficePath
        }
    }

    if ([string]::IsNullOrWhiteSpace($OfficePath)) { throw 'OfficePath fehlt. Nutze -Interactive oder uebergib -OfficePath.' }
    if ([string]::IsNullOrWhiteSpace($CustomUiPath)) { throw 'CustomUiPath fehlt. Nutze -Interactive oder uebergib -CustomUiPath.' }

    $resolvedOfficePath = Resolve-RequiredFile -Path $OfficePath -Name 'OfficePath'
    $resolvedCustomUiPath = Resolve-RequiredFile -Path $CustomUiPath -Name 'CustomUiPath'
    Assert-SupportedOfficePath -Path $resolvedOfficePath -Name 'OfficePath'

    if ($ForceInPlace -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
        throw 'OutputPath darf nicht zusammen mit ForceInPlace verwendet werden.'
    }

    if ($ForceInPlace) {
        $targetPath = $resolvedOfficePath
    }
    else {
        if ([string]::IsNullOrWhiteSpace($OutputPath)) {
            $sourceDirectory = [System.IO.Path]::GetDirectoryName($resolvedOfficePath)
            $sourceBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedOfficePath)
            $sourceExtension = [System.IO.Path]::GetExtension($resolvedOfficePath)
            $targetPath = Join-Path $sourceDirectory "$sourceBaseName.customui$sourceExtension"
        }
        else {
            $targetPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
        }
        Assert-SupportedOfficePath -Path $targetPath -Name 'OutputPath'
        if ([string]::Equals($resolvedOfficePath, $targetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'OutputPath entspricht OfficePath. Verwende ForceInPlace, wenn die Originaldatei ueberschrieben werden soll.'
        }
        $targetDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
        if (-not (Test-Path -LiteralPath $targetDirectory -PathType Container)) {
            throw "OutputPath-Verzeichnis wurde nicht gefunden: $targetDirectory"
        }
        Copy-Item -LiteralPath $resolvedOfficePath -Destination $targetPath -Force
    }

    $xmlText = Get-Content -Raw -LiteralPath $resolvedCustomUiPath
    Assert-CustomUiXml -XmlText $xmlText
    Set-CustomUi14Xml -TargetPath $targetPath -XmlText $xmlText

    Write-Host 'CustomUI14 XML wurde eingespielt.'
    Write-Host "OfficePath: $resolvedOfficePath"
    Write-Host "CustomUiPath: $resolvedCustomUiPath"
    Write-Host "OutputPath: $targetPath"
    exit 0
}
catch {
    $message = $_.Exception.Message
    if ($_.Exception -is [System.IO.IOException] -or $_.Exception.InnerException -is [System.IO.IOException]) {
        $message = "$message Bitte pruefen, ob die Office-Datei geoeffnet ist, und Office schliessen."
    }
    [Console]::Error.WriteLine("ERROR: $message")
    exit 1
}

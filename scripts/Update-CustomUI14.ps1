[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string] $DefinitionPath,

    [string] $OutputPath,

    [switch] $ForceInPlace
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$CustomUi14Namespace = 'http://schemas.microsoft.com/office/2009/07/customui'
$CustomUiRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'

function Resolve-RequiredFile {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Name wurde nicht gefunden: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).ProviderPath
}

function Assert-SupportedOfficePath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    $extension = [System.IO.Path]::GetExtension($Path)
    if ($extension -ine '.xlsm' -and $extension -ine '.xlam') {
        throw "$Name muss eine .xlsm- oder .xlam-Datei sein: $Path"
    }
}

function Split-DefinitionLine {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Line,

        [Parameter(Mandatory = $true)]
        [int] $LineNumber
    )

    $fields = $Line.Split(';')
    for ($i = 0; $i -lt $fields.Count; $i++) {
        $fields[$i] = $fields[$i].Trim()
    }

    if ($fields.Count -ne 5) {
        throw "Definition Zeile ${LineNumber}: Erwartet 5 ;-getrennte Felder, gefunden $($fields.Count)."
    }

    return $fields
}

function Read-RibbonDefinition {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $requiredHeaders = @('tab_id', 'tab_label', 'group_id', 'group_label')
    $metadata = @{}
    $buttons = @()
    $seenIds = @{}
    $tableStarted = $false
    $expectedHeader = @('id', 'label', 'size', 'icon', 'onAction')
    $lines = Get-Content -LiteralPath $Path

    for ($lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
        $lineNumber = $lineIndex + 1
        $line = $lines[$lineIndex].Trim()

        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if ($line.StartsWith('#')) {
            if ($tableStarted) {
                continue
            }

            $metadataMatch = [regex]::Match($line, '^#(?<key>[^=]+)=(?<value>.*)$')
            if (-not $metadataMatch.Success) {
                continue
            }

            $key = $metadataMatch.Groups['key'].Value.Trim()
            $value = $metadataMatch.Groups['value'].Value.Trim()
            if ([string]::IsNullOrWhiteSpace($key) -or [string]::IsNullOrWhiteSpace($value)) {
                throw "Definition Zeile ${lineNumber}: Ungueltige Metadatenzeile."
            }

            $metadata[$key] = $value
            continue
        }

        $fields = Split-DefinitionLine -Line $line -LineNumber $lineNumber

        if (-not $tableStarted) {
            for ($i = 0; $i -lt $expectedHeader.Count; $i++) {
                if ($fields[$i] -cne $expectedHeader[$i]) {
                    throw "Definition Zeile ${lineNumber}: Erwarteter Tabellenkopf ist 'id;label;size;icon;onAction'."
                }
            }

            $tableStarted = $true
            continue
        }

        $id = $fields[0]
        $label = $fields[1]
        $size = $fields[2]
        $icon = $fields[3]
        $onAction = $fields[4]

        if ([string]::IsNullOrWhiteSpace($id)) {
            throw "Definition Zeile ${lineNumber}: id ist Pflicht."
        }

        if ([string]::IsNullOrWhiteSpace($label)) {
            throw "Definition Zeile ${lineNumber}: label ist Pflicht."
        }

        if ($size -cne 'normal' -and $size -cne 'large') {
            throw "Definition Zeile ${lineNumber}: size muss 'normal' oder 'large' sein."
        }

        if ([string]::IsNullOrWhiteSpace($icon)) {
            throw "Definition Zeile ${lineNumber}: icon ist Pflicht."
        }

        if ([string]::IsNullOrWhiteSpace($onAction)) {
            throw "Definition Zeile ${lineNumber}: onAction ist Pflicht."
        }

        if ($seenIds.ContainsKey($id)) {
            throw "Definition Zeile ${lineNumber}: Doppelte id '$id' (bereits in Zeile $($seenIds[$id]))."
        }

        $seenIds[$id] = $lineNumber
        $buttons += [pscustomobject] @{
            Id = $id
            Label = $label
            Size = $size
            Icon = $icon
            OnAction = $onAction
            LineNumber = $lineNumber
        }
    }

    if (-not $tableStarted) {
        throw "Definition enthaelt keinen Tabellenkopf 'id;label;size;icon;onAction'."
    }

    foreach ($headerName in $requiredHeaders) {
        if (-not $metadata.ContainsKey($headerName) -or [string]::IsNullOrWhiteSpace($metadata[$headerName])) {
            throw "Definition fehlt Pflicht-Metadatum #$headerName=..."
        }
    }

    return [pscustomobject] @{
        TabId = $metadata['tab_id']
        TabLabel = $metadata['tab_label']
        GroupId = $metadata['group_id']
        GroupLabel = $metadata['group_label']
        InsertAfterMso = $metadata['insert_after_mso']
        Buttons = $buttons
    }
}

function Get-DirectElement {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode] $Parent,

        [Parameter(Mandatory = $true)]
        [string] $LocalName,

        [string] $AttributeName,

        [string] $AttributeValue
    )

    foreach ($child in $Parent.ChildNodes) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) {
            continue
        }

        if ($child.LocalName -ne $LocalName -or $child.NamespaceURI -ne $CustomUi14Namespace) {
            continue
        }

        if ([string]::IsNullOrEmpty($AttributeName)) {
            return $child
        }

        if ($child.GetAttribute($AttributeName) -ceq $AttributeValue) {
            return $child
        }
    }

    return $null
}

function New-CustomUiElement {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument] $Document,

        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    return $Document.CreateElement($Name, $CustomUi14Namespace)
}

function Get-OrCreateChildElement {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode] $Parent,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument] $Document,

        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    $child = Get-DirectElement -Parent $Parent -LocalName $Name
    if ($null -ne $child) {
        return $child
    }

    $child = New-CustomUiElement -Document $Document -Name $Name
    [void] $Parent.AppendChild($child)
    return $child
}

function Get-CustomUi14Relationship {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Packaging.Package] $Package
    )

    foreach ($relationship in $Package.GetRelationshipsByType($CustomUiRelationshipType)) {
        if ($relationship.TargetMode -ne [System.IO.Packaging.TargetMode]::Internal) {
            continue
        }

        $target = $relationship.TargetUri.OriginalString.TrimStart('/')
        if ($target -ieq 'customUI/customUI14.xml') {
            return $relationship
        }
    }

    return $null
}

function Read-CustomUiDocument {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Packaging.PackagePart] $Part
    )

    $document = New-Object System.Xml.XmlDocument
    $document.PreserveWhitespace = $false

    $stream = $null
    try {
        $stream = $Part.GetStream([System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        if ($stream.Length -eq 0) {
            $document.LoadXml('<customUI xmlns="' + $CustomUi14Namespace + '"><ribbon><tabs /></ribbon></customUI>')
        } else {
            $document.Load($stream)
        }
    } finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }

    if ($document.DocumentElement.LocalName -ne 'customUI' -or $document.DocumentElement.NamespaceURI -ne $CustomUi14Namespace) {
        throw "Vorhandenes customUI14.xml verwendet nicht den erwarteten Namespace $CustomUi14Namespace."
    }

    return $document
}

function Save-CustomUiDocument {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument] $Document,

        [Parameter(Mandatory = $true)]
        [System.IO.Packaging.PackagePart] $Part
    )

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)

    $stream = $null
    $writer = $null
    try {
        $stream = $Part.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        $Document.Save($writer)
    } finally {
        if ($null -ne $writer) {
            $writer.Dispose()
        } elseif ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Update-CustomUi14 {
    param(
        [Parameter(Mandatory = $true)]
        [string] $TargetPath,

        [Parameter(Mandatory = $true)]
        [pscustomobject] $Definition
    )

    Add-Type -AssemblyName WindowsBase

    $partUri = New-Object System.Uri('/customUI/customUI14.xml', [System.UriKind]::Relative)
    $relationshipUri = New-Object System.Uri('customUI/customUI14.xml', [System.UriKind]::Relative)
    $package = $null
    $createdPart = $false
    $createdTabs = 0
    $createdGroups = 0
    $createdButtons = 0
    $updatedButtons = 0
    $deletedButtons = 0

    try {
        $package = [System.IO.Packaging.Package]::Open($TargetPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
        $relationship = Get-CustomUi14Relationship -Package $package
        if ($null -eq $relationship) {
            [void] $package.CreateRelationship($relationshipUri, [System.IO.Packaging.TargetMode]::Internal, $CustomUiRelationshipType)
            Write-Verbose "Package-Relationship fuer customUI14.xml erstellt."
        }

        if ($package.PartExists($partUri)) {
            $part = $package.GetPart($partUri)
        } else {
            $part = $package.CreatePart($partUri, 'application/xml', [System.IO.Packaging.CompressionOption]::Maximum)
            $createdPart = $true
            Write-Verbose "Part /customUI/customUI14.xml erstellt."
        }

        $document = Read-CustomUiDocument -Part $part
        $customUi = $document.DocumentElement
        $ribbon = Get-OrCreateChildElement -Parent $customUi -Document $document -Name 'ribbon'
        $tabs = Get-OrCreateChildElement -Parent $ribbon -Document $document -Name 'tabs'

        $tab = Get-DirectElement -Parent $tabs -LocalName 'tab' -AttributeName 'id' -AttributeValue $Definition.TabId
        if ($null -eq $tab) {
            $tab = New-CustomUiElement -Document $document -Name 'tab'
            [void] $tabs.AppendChild($tab)
            $createdTabs++
        }

        $tab.SetAttribute('id', $Definition.TabId)
        $tab.SetAttribute('label', $Definition.TabLabel)
        if (-not [string]::IsNullOrWhiteSpace($Definition.InsertAfterMso)) {
            $tab.SetAttribute('insertAfterMso', $Definition.InsertAfterMso)
        }

        $group = Get-DirectElement -Parent $tab -LocalName 'group' -AttributeName 'id' -AttributeValue $Definition.GroupId
        if ($null -eq $group) {
            $group = New-CustomUiElement -Document $document -Name 'group'
            [void] $tab.AppendChild($group)
            $createdGroups++
        }

        $group.SetAttribute('id', $Definition.GroupId)
        $group.SetAttribute('label', $Definition.GroupLabel)

        $definitionButtonIds = @{}
        foreach ($buttonDefinition in $Definition.Buttons) {
            $definitionButtonIds[$buttonDefinition.Id] = $true
        }

        $buttonsToDelete = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
        foreach ($child in $group.ChildNodes) {
            if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) {
                continue
            }

            if ($child.LocalName -ne 'button' -or $child.NamespaceURI -ne $CustomUi14Namespace) {
                continue
            }

            $existingButtonId = $child.GetAttribute('id')
            if (-not [string]::IsNullOrWhiteSpace($existingButtonId) -and $definitionButtonIds.ContainsKey($existingButtonId)) {
                continue
            }

            [void] $buttonsToDelete.Add($child)
        }

        foreach ($buttonToDelete in $buttonsToDelete) {
            [void] $group.RemoveChild($buttonToDelete)
            $deletedButtons++
        }

        foreach ($buttonDefinition in $Definition.Buttons) {
            $button = Get-DirectElement -Parent $group -LocalName 'button' -AttributeName 'id' -AttributeValue $buttonDefinition.Id
            if ($null -eq $button) {
                $button = New-CustomUiElement -Document $document -Name 'button'
                [void] $group.AppendChild($button)
                $createdButtons++
            } else {
                $updatedButtons++
            }

            $button.SetAttribute('id', $buttonDefinition.Id)
            $button.SetAttribute('label', $buttonDefinition.Label)
            $button.SetAttribute('size', $buttonDefinition.Size)
            $button.SetAttribute('imageMso', $buttonDefinition.Icon)
            $button.SetAttribute('onAction', $buttonDefinition.OnAction)
        }

        Save-CustomUiDocument -Document $document -Part $part
        $package.Flush()
    } finally {
        if ($null -ne $package) {
            $package.Close()
        }
    }

    return [pscustomobject] @{
        CreatedPart = $createdPart
        CreatedTabs = $createdTabs
        CreatedGroups = $createdGroups
        CreatedButtons = $createdButtons
        UpdatedButtons = $updatedButtons
        DeletedButtons = $deletedButtons
        OutputPath = $TargetPath
    }
}

try {
    $resolvedWorkbookPath = Resolve-RequiredFile -Path $WorkbookPath -Name 'WorkbookPath'
    $resolvedDefinitionPath = Resolve-RequiredFile -Path $DefinitionPath -Name 'DefinitionPath'
    Assert-SupportedOfficePath -Path $resolvedWorkbookPath -Name 'WorkbookPath'

    if ($ForceInPlace -and -not [string]::IsNullOrWhiteSpace($OutputPath)) {
        throw "OutputPath darf nicht zusammen mit ForceInPlace verwendet werden."
    }

    if ($ForceInPlace) {
        $targetPath = $resolvedWorkbookPath
    } else {
        if ([string]::IsNullOrWhiteSpace($OutputPath)) {
            $sourceDirectory = [System.IO.Path]::GetDirectoryName($resolvedWorkbookPath)
            $sourceBaseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedWorkbookPath)
            $sourceExtension = [System.IO.Path]::GetExtension($resolvedWorkbookPath)
            $targetPath = Join-Path $sourceDirectory "$sourceBaseName.ribbon$sourceExtension"
        } else {
            $targetPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
        }

        Assert-SupportedOfficePath -Path $targetPath -Name 'OutputPath'

        if ([string]::Equals($resolvedWorkbookPath, $targetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "OutputPath entspricht WorkbookPath. Verwende ForceInPlace, wenn die Originaldatei ueberschrieben werden soll."
        }

        $targetDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
        if (-not (Test-Path -LiteralPath $targetDirectory -PathType Container)) {
            throw "OutputPath-Verzeichnis wurde nicht gefunden: $targetDirectory"
        }

        Copy-Item -LiteralPath $resolvedWorkbookPath -Destination $targetPath -Force
    }

    $definition = Read-RibbonDefinition -Path $resolvedDefinitionPath
    $result = Update-CustomUi14 -TargetPath $targetPath -Definition $definition

    Write-Host "CustomUI14 wurde aktualisiert."
    Write-Host "CreatedTabs: $($result.CreatedTabs)"
    Write-Host "CreatedGroups: $($result.CreatedGroups)"
    Write-Host "CreatedButtons: $($result.CreatedButtons)"
    Write-Host "UpdatedButtons: $($result.UpdatedButtons)"
    Write-Host "DeletedButtons: $($result.DeletedButtons)"
    Write-Host "OutputPath: $($result.OutputPath)"
    exit 0
} catch {
    $message = $_.Exception.Message
    if ($_.Exception -is [System.IO.IOException] -or $_.Exception.InnerException -is [System.IO.IOException]) {
        $message = "$message Bitte pruefen, ob die Datei in Excel geoeffnet ist, und Excel schliessen."
    }

    [Console]::Error.WriteLine("ERROR: $message")
    exit 1
}

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $WorkbookPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$CustomUi14Namespace = 'http://schemas.microsoft.com/office/2009/07/customui'
$CustomUiRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'

function Get-CustomUi14Relationship {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Packaging.Package] $Package
    )

    foreach ($relationship in $Package.GetRelationshipsByType($CustomUiRelationshipType)) {
        if ($relationship.TargetMode -ne [System.IO.Packaging.TargetMode]::Internal) {
            continue
        }

        if ($relationship.TargetUri.OriginalString.TrimStart('/') -ieq 'customUI/customUI14.xml') {
            return $relationship
        }
    }

    return $null
}

try {
    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
        throw "WorkbookPath wurde nicht gefunden: $WorkbookPath"
    }

    Add-Type -AssemblyName WindowsBase

    $resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).ProviderPath
    $partUri = New-Object System.Uri('/customUI/customUI14.xml', [System.UriKind]::Relative)
    $package = $null

    try {
        $package = [System.IO.Packaging.Package]::Open($resolvedWorkbookPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $relationship = Get-CustomUi14Relationship -Package $package
        if ($null -eq $relationship -or -not $package.PartExists($partUri)) {
            Write-Host "customUI14.xml: nicht vorhanden"
            exit 2
        }

        $part = $package.GetPart($partUri)
        $document = New-Object System.Xml.XmlDocument
        $stream = $part.GetStream([System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        try {
            $document.Load($stream)
        } finally {
            $stream.Dispose()
        }

        if ($document.DocumentElement.NamespaceURI -ne $CustomUi14Namespace) {
            throw "customUI14.xml verwendet nicht den erwarteten Namespace $CustomUi14Namespace."
        }

        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
        $namespaceManager.AddNamespace('ui', $CustomUi14Namespace)
        $buttons = $document.SelectNodes('/ui:customUI/ui:ribbon/ui:tabs/ui:tab/ui:group/ui:button', $namespaceManager)

        Write-Host "customUI14.xml: vorhanden"
        Write-Host "Namespace: $($document.DocumentElement.NamespaceURI)"
        foreach ($button in $buttons) {
            [pscustomobject] @{
                TabId = $button.ParentNode.ParentNode.GetAttribute('id')
                GroupId = $button.ParentNode.GetAttribute('id')
                Id = $button.GetAttribute('id')
                Label = $button.GetAttribute('label')
                Size = $button.GetAttribute('size')
                ImageMso = $button.GetAttribute('imageMso')
                OnAction = $button.GetAttribute('onAction')
            }
        }
    } finally {
        if ($null -ne $package) {
            $package.Close()
        }
    }
} catch {
    [Console]::Error.WriteLine("ERROR: $($_.Exception.Message)")
    exit 1
}

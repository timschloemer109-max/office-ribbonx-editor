[CmdletBinding()]
param(
    [string] $OfficePath,
    [string] $OutputPath,
    [switch] $ForceInPlace,
    [ValidateSet('12', '14', 'all')]
    [string] $Type = 'all'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$CustomUi12RelationshipType = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
$CustomUi14RelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'

function Assert-SupportedOfficePath {
    param([string] $Path, [string] $Name)
    $supportedExtensions = @('.xlsm', '.xlam', '.xltm', '.docm', '.dotm', '.pptm', '.ppam')
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

function Remove-EntryIfPresent {
    param([System.IO.Compression.ZipArchive] $Zip, [string] $EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if ($null -ne $entry) { $entry.Delete(); return $true }
    return $false
}

function Remove-CustomUiRelationships {
    param([System.IO.Compression.ZipArchive] $Zip, [string] $Mode)

    $relsText = Read-ZipText -Zip $Zip -EntryName '_rels/.rels'
    if ([string]::IsNullOrWhiteSpace($relsText)) { return 0 }

    [xml]$document = $relsText
    $relationships = @($document.Relationships.Relationship)
    $toRemove = @()

    foreach ($relationship in $relationships) {
        $targetValue = ''
        if ($null -ne $relationship.Target) { $targetValue = [string]$relationship.Target }
        $target = $targetValue.TrimStart('/').ToLowerInvariant()
        $removeFor12 = $Mode -in @('12', 'all') -and $relationship.Type -eq $CustomUi12RelationshipType -and $target -eq 'customui/customui.xml'
        $removeFor14 = $Mode -in @('14', 'all') -and $relationship.Type -eq $CustomUi14RelationshipType -and $target -eq 'customui/customui14.xml'
        if ($removeFor12 -or $removeFor14) {
            $toRemove += $relationship
        }
    }

    foreach ($relationship in $toRemove) {
        [void]$document.Relationships.RemoveChild($relationship)
    }

    if ($toRemove.Count -gt 0) {
        Write-ZipText -Zip $Zip -EntryName '_rels/.rels' -Text $document.OuterXml
    }

    return $toRemove.Count
}

function Remove-CustomUiXml {
    param([string] $TargetPath, [string] $Mode)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::Open($TargetPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $deletedCustomUi = 0
        if ($Mode -in @('12', 'all')) {
            if (Remove-EntryIfPresent -Zip $zip -EntryName 'customUI/customUI.xml') { $deletedCustomUi++ }
        }
        if ($Mode -in @('14', 'all')) {
            if (Remove-EntryIfPresent -Zip $zip -EntryName 'customUI/customUI14.xml') { $deletedCustomUi++ }
        }

        $deletedRelationships = Remove-CustomUiRelationships -Zip $zip -Mode $Mode

        Write-Host "DeletedCustomUiParts: $deletedCustomUi"
        Write-Host "DeletedRelationships: $deletedRelationships"
    }
    finally {
        $zip.Dispose()
    }
}

try {
    if ([string]::IsNullOrWhiteSpace($OfficePath)) { throw 'OfficePath fehlt.' }

    $resolvedOfficePath = Resolve-RequiredFile -Path $OfficePath -Name 'OfficePath'
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
            $targetPath = Join-Path $sourceDirectory "$sourceBaseName.nocustomui$sourceExtension"
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

    Remove-CustomUiXml -TargetPath $targetPath -Mode $Type

    Write-Host 'CustomUI wurde entfernt.'
    Write-Host "OfficePath: $resolvedOfficePath"
    Write-Host "OutputPath: $targetPath"
    Write-Host "Type: $Type"
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

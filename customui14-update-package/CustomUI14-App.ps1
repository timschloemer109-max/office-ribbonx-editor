[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$setScript = Join-Path $scriptRoot 'Set-CustomUI14Xml.ps1'
$removeScript = Join-Path $scriptRoot 'Remove-CustomUIXml.ps1'

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 130)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, 24)
    return $label
}

function New-TextBox {
    param([int]$X, [int]$Y)
    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point($X, $Y)
    $box.Size = New-Object System.Drawing.Size(440, 24)
    $box.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    return $box
}

function New-Button {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 135)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, 28)
    return $button
}

function Select-OfficeFileForBox {
    param([System.Windows.Forms.TextBox]$Target)
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'DOF-/Office-Datei auswaehlen'
    $dialog.Filter = 'Office Makrodateien (*.xlsm;*.xlam;*.xltm;*.docm;*.dotm;*.pptm;*.ppam)|*.xlsm;*.xlam;*.xltm;*.docm;*.dotm;*.pptm;*.ppam|Alle Dateien (*.*)|*.*'
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Target.Text = $dialog.FileName
    }
}

function Select-CustomUiFileForBox {
    param([System.Windows.Forms.TextBox]$Target)
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'customUI14.txt auswaehlen'
    $dialog.Filter = 'CustomUI14 XML (*.txt;*.xml)|*.txt;*.xml|Alle Dateien (*.*)|*.*'
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Target.Text = $dialog.FileName
    }
}

function Select-OutputFileForBox {
    param(
        [System.Windows.Forms.TextBox]$OfficeBox,
        [System.Windows.Forms.TextBox]$OutputBox
    )
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = 'Ausgabedatei speichern'
    $dialog.Filter = 'Office Makrodateien (*.xlsm;*.xlam;*.xltm;*.docm;*.dotm;*.pptm;*.ppam)|*.xlsm;*.xlam;*.xltm;*.docm;*.dotm;*.pptm;*.ppam|Alle Dateien (*.*)|*.*'
    if (-not [string]::IsNullOrWhiteSpace($OfficeBox.Text)) {
        $dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($OfficeBox.Text)
        $dialog.FileName = [System.IO.Path]::GetFileNameWithoutExtension($OfficeBox.Text) + '.customui' + [System.IO.Path]::GetExtension($OfficeBox.Text)
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $OutputBox.Text = $dialog.FileName
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'CustomUI in DOF-Datei verwalten'
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Size = New-Object System.Drawing.Size(850, 410)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(820, 400)

$modeLabel = New-Label 'Modus' 16 24
$modeCombo = New-Object System.Windows.Forms.ComboBox
$modeCombo.Location = New-Object System.Drawing.Point(160, 22)
$modeCombo.Size = New-Object System.Drawing.Size(220, 24)
$modeCombo.DropDownStyle = 'DropDownList'
[void]$modeCombo.Items.Add('CustomUI einspielen')
[void]$modeCombo.Items.Add('CustomUI loeschen')
$modeCombo.SelectedIndex = 0

$removeTypeLabel = New-Label 'Loesch-Typ' 410 24 90
$removeTypeCombo = New-Object System.Windows.Forms.ComboBox
$removeTypeCombo.Location = New-Object System.Drawing.Point(510, 22)
$removeTypeCombo.Size = New-Object System.Drawing.Size(140, 24)
$removeTypeCombo.DropDownStyle = 'DropDownList'
[void]$removeTypeCombo.Items.Add('all')
[void]$removeTypeCombo.Items.Add('12')
[void]$removeTypeCombo.Items.Add('14')
$removeTypeCombo.SelectedIndex = 0

$officeLabel = New-Label '1. DOF-/Office-Datei' 16 64
$officeBox = New-TextBox 160 62
$officeButton = New-Button 'Durchsuchen...' 680 60
$officeButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$officeButton.Add_Click({ Select-OfficeFileForBox -Target $officeBox })

$customLabel = New-Label '2. customUI14' 16 108
$customBox = New-TextBox 160 106
$customButton = New-Button 'Durchsuchen...' 680 104
$customButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$customButton.Add_Click({ Select-CustomUiFileForBox -Target $customBox })

$outputLabel = New-Label 'Ausgabedatei' 16 152
$outputBox = New-TextBox 160 150
$outputButton = New-Button 'Speichern...' 680 148
$outputButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$outputButton.Add_Click({ Select-OutputFileForBox -OfficeBox $officeBox -OutputBox $outputBox })

$forceCheck = New-Object System.Windows.Forms.CheckBox
$forceCheck.Text = 'Originaldatei direkt ueberschreiben (Office vorher schliessen)'
$forceCheck.Location = New-Object System.Drawing.Point(160, 188)
$forceCheck.Size = New-Object System.Drawing.Size(520, 24)
$forceCheck.Add_CheckedChanged({
    $outputBox.Enabled = -not $forceCheck.Checked
    $outputButton.Enabled = -not $forceCheck.Checked
})

$runButton = New-Button 'Ausfuehren' 160 226 180
$closeButton = New-Button 'Schliessen' 350 226 120

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(16, 268)
$logBox.Size = New-Object System.Drawing.Size(800, 98)
$logBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$logBox.Multiline = $true
$logBox.ScrollBars = 'Vertical'
$logBox.ReadOnly = $true

$refreshModeUI = {
    $removeMode = $modeCombo.SelectedItem -eq 'CustomUI loeschen'
    $customLabel.Enabled = -not $removeMode
    $customBox.Enabled = -not $removeMode
    $customButton.Enabled = -not $removeMode
    $removeTypeLabel.Enabled = $removeMode
    $removeTypeCombo.Enabled = $removeMode
    if ($removeMode) {
        $runButton.Text = 'CustomUI loeschen'
    }
    else {
        $runButton.Text = 'CustomUI einspielen'
    }
}
$modeCombo.Add_SelectedIndexChanged($refreshModeUI)
& $refreshModeUI

$runButton.Add_Click({
    $logBox.Clear()

    if ([string]::IsNullOrWhiteSpace($officeBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Bitte zuerst die DOF-/Office-Datei auswaehlen.', 'Fehlt noch', 'OK', 'Warning') | Out-Null
        return
    }

    $removeMode = $modeCombo.SelectedItem -eq 'CustomUI loeschen'

    if (-not $removeMode -and [string]::IsNullOrWhiteSpace($customBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Bitte die customUI14.txt oder customUI14.xml auswaehlen.', 'Fehlt noch', 'OK', 'Warning') | Out-Null
        return
    }

    if (-not $forceCheck.Checked -and [string]::IsNullOrWhiteSpace($outputBox.Text)) {
        $suffix = if ($removeMode) { '.nocustomui' } else { '.customui' }
        $defaultOutput = [System.IO.Path]::Combine(
            [System.IO.Path]::GetDirectoryName($officeBox.Text),
            [System.IO.Path]::GetFileNameWithoutExtension($officeBox.Text) + $suffix + [System.IO.Path]::GetExtension($officeBox.Text)
        )
        $outputBox.Text = $defaultOutput
    }

    if ($removeMode) {
        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $removeScript,
            '-OfficePath', $officeBox.Text,
            '-Type', $removeTypeCombo.SelectedItem
        )
    }
    else {
        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $setScript,
            '-OfficePath', $officeBox.Text,
            '-CustomUiPath', $customBox.Text
        )
    }

    if ($forceCheck.Checked) {
        $args += '-ForceInPlace'
    }
    else {
        $args += @('-OutputPath', $outputBox.Text)
    }

    $runButton.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $output = & powershell.exe @args 2>&1
        $logBox.Text = ($output -join [Environment]::NewLine)
        if ($LASTEXITCODE -eq 0) {
            $doneMessage = if ($removeMode) { 'CustomUI wurde erfolgreich geloescht.' } else { 'CustomUI wurde erfolgreich eingespielt.' }
            [System.Windows.Forms.MessageBox]::Show($doneMessage, 'Fertig', 'OK', 'Information') | Out-Null
        }
        else {
            [System.Windows.Forms.MessageBox]::Show('Die Ausfuehrung ist fehlgeschlagen. Details stehen im Log.', 'Fehler', 'OK', 'Error') | Out-Null
        }
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $runButton.Enabled = $true
    }
})

$closeButton.Add_Click({ $form.Close() })

$form.Controls.AddRange(@(
    $modeLabel, $modeCombo, $removeTypeLabel, $removeTypeCombo,
    $officeLabel, $officeBox, $officeButton,
    $customLabel, $customBox, $customButton,
    $outputLabel, $outputBox, $outputButton,
    $forceCheck, $runButton, $closeButton, $logBox
))

[void] $form.ShowDialog()

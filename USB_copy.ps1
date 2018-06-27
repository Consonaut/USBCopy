# Author: Constantin Heine <it@staedelschule.de> 2018
# Copy connected USB drives to a set destination
#
#

#Function to hide the Powershell terminal - https://stackoverflow.com/a/46414223
$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
Function Show-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 10)
}
Function Hide-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}

#Hide the Powershell Terminal
Hide-Powershell

#Add Forms support
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set and create Destination location
$destination = "c:\temp\WS2019\"
If ((test-path $destination) -eq $FALSE) {
    mkdir -Path $destination > $null
}

#Event code for clicking $CopyButton
$CopyButton_OnClick = {
    $textbox.focus() #Refocus the textbox
    $ApplicantNumber = $textBox.Text
    $ApplicantFolder = $destination + $ApplicantNumber #The copy destination
    $sourceDrive = $dropdown.SelectedItem #The copy source
    $FindDrives = Get-WmiObject -Class win32_logicaldisk | where-object {($_.DriveType -eq '2') -and ($_.Size -ne $null)} | Select-Object -ExpandProperty DeviceID
    If ($FindDrives -contains $sourceDrive) {
        If ((test-path $ApplicantFolder) -eq $FALSE) {
            mkdir -Path $ApplicantFolder > $null #Create the destination if necessary
            $process = Start-Process robocopy.exe -windowstyle Hidden -ArgumentList $sourceDrive, $ApplicantFolder, '/np /s /copy:DT /dcopy:DT' -PassThru -Wait
            $exitcode = $process.ExitCode
            If ($exitcode -gt 1) { [System.Windows.Forms.MessageBox]::Show("Fehler!","Fehler beim Kopieren!",0) } #Catch exit codes > 1
            Else {
                $Eject = New-Object -comObject Shell.Application
                $Eject.NameSpace(17).ParseName($sourceDrive).InvokeVerb(“Eject”) #Eject the drive if the copy was successful
                [System.Windows.Forms.MessageBox]::Show("Laufwerk " + $sourceDrive + " nach " + $ApplicantFolder + " kopiert.","Fertig.",0) }}
        Else { $result = [System.Windows.Forms.Messagebox]::Show("Dieser Bewerber existiert bereits, klicke Wiederholen zum Überschreiben oder Abbrechen zum beenden.","Fehler beim Kopieren!",5) }
        } 
        Else { [System.Windows.Forms.MessageBox]::Show(„Quelllaufwerk nicht gefunden.“,“Fehler beim Lesen von USB!“,0) }
            If ($result -eq "Retry") {
                $FindDrives = Get-WmiObject -Class win32_logicaldisk | where-object {($_.DriveType -eq '2') -and ($_.Size -ne $null)} | Select-Object -ExpandProperty DeviceID
                If ($FindDrives -contains $sourceDrive) {
                    $process = Start-Process robocopy.exe -windowstyle Hidden -ArgumentList $sourceDrive, $ApplicantFolder, '/np /s /copy:DT /dcopy:DT' -PassThru -Wait
                    $exitcode = $process.ExitCode
                    If ($exitcode -gt 1) { [System.Windows.Forms.MessageBox]::Show("Fehler!","Fehler beim Kopieren!",0) } #Catch exit codes > 1
                    Else {
                        $Eject = New-Object -comObject Shell.Application
                        $Eject.NameSpace(17).ParseName($sourceDrive).InvokeVerb(“Eject”) #Eject the drive if the copy was successful
                        [System.Windows.Forms.MessageBox]::Show("Laufwerk " + $sourceDrive + " nach " + $ApplicantFolder + " kopiert.","Fertig.",0) }
                }
                Else { [System.Windows.Forms.MessageBox]::Show(„Quelllaufwerk nicht gefunden.“,“Fehler beim Lesen von USB!“,0) }
                    }
            Else {
                $a = "Schreibvorgang wurde abgebrochen"
                $a }
}


#Event code for clicking $RefreshButton
$RefreshButton_OnClick = {
    $textBox.focus()
    $dropdown.Items.Clear()
    $FindDrives = Get-WmiObject -Class win32_logicaldisk | where-object {($_.DriveType -eq '2') -and ($_.Size -ne $null)} | Select-Object -ExpandProperty DeviceID
    Foreach ($drive in $FindDrives) { $dropdown.Items.Add($drive) }
    $dropdown.SelectedItem = $dropdown.Items[0]
}

#Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'USB Copy'
$form.Size = '465,300'
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font("Segoe UI",13,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = '340,200'
$CancelButton.Size = '100,50'
$CancelButton.Text = 'Abbrechen'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$CopyButton = New-Object System.Windows.Forms.Button
$CopyButton.Location = '10,200'
$CopyButton.Size = '100,50'
$CopyButton.Text = 'Kopieren'
$CopyButton.add_click($CopyButton_OnClick)
$form.Controls.Add($CopyButton)
$form.AcceptButton = $CopyButton

$RefreshButton = New-Object System.Windows.Forms.Button
$RefreshButton.Location = '175,200'
$RefreshButton.Size = '100,50'
$RefreshButton.Text = 'Aktualisieren'
$RefreshButton.add_click($RefreshButton_OnClick)
$form.Controls.Add($RefreshButton)

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = '10,10'
$label1.Size = '440,60'
$label1.Text = 'USB Laufwerk anschließen und Bewerber Nummer eingeben. Klicke Kopieren um zu Starten, Aktualisieren um USB Laufwerke neu einzulesen, Abbrechen um das Programm zu beenden.'
$form.Controls.Add($label1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = '10,87'
$label2.Size = '130,20'
$label2.Text = 'Bewerber Number:'
$form.Controls.Add($label2)

$labelDestination = New-Object System.Windows.Forms.Label
$labelDestination.Location = '10,147'
$labelDestination.Size = '130,20'
$labelDestination.Text = 'Ziel Verzeichniss:'
$form.Controls.Add($labelDestination)

$labelDropdown = New-Object System.Windows.Forms.Label
$labelDropdown.Location = '10,119'
$labelDropdown.Size = '130,20'
$labelDropdown.Text = 'Quell USB Laufwerk:'
$form.Controls.Add($labelDropdown)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = '140,85'
$textBox.Size = '200,20'
$textBox.ShortcutsEnabled = 'True'
$form.Controls.Add($textBox)

$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = '140,115'
$dropdown.Size = '200,20'
$FindDrives = Get-WmiObject -Class win32_logicaldisk | where-object {($_.DriveType -eq '2') -and ($_.Size -ne $null)} | Select-Object -ExpandProperty DeviceID
Foreach ($drive in $FindDrives) { $dropdown.Items.Add($drive) }
$dropdown.SelectedItem = $dropdown.Items[0]
$form.Controls.Add($dropdown)

$textBoxDestination = New-Object System.Windows.Forms.TextBox
$textBoxDestination.Location = '140,145'
$textBoxDestination.Size = '200,20'
$textBoxDestination.Text = $destination
$textBoxDestination.ShortcutsEnabled = 'True'
$form.Controls.Add($textBoxDestination)

$form.Add_Shown({$form.Activate();$textbox.focus()})
$form.KeyPreview = $true #Necessary to enable pressing Enter in the textbox
$form.Topmost = $true
$form.ShowDialog()


# Retrieve all VMs, their hardware versions, and the host server they are assigned to
$allVMs = Get-VM | Select-Object Name, 
                                      @{N="HardwareVersion"; E={$_.ExtensionData.Config.Version}}, 
                                      @{N="Host"; E={$_.VMHost.Name}}

$uniqueVersions = $allVMs | Select-Object -ExpandProperty HardwareVersion -Unique

# Create a Windows Form to display a dropdown of unique hardware versions
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text = "Select a VM Hardware Version"
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = "Please select a hardware version:"
$form.Controls.Add($label)

$dropdown = New-Object System.Windows.Forms.ComboBox
$dropdown.Location = New-Object System.Drawing.Point(10,40)
$dropdown.Size = New-Object System.Drawing.Size(260,30)
foreach ($version in $uniqueVersions) {
    $dropdown.Items.Add($version)
}
$form.Controls.Add($dropdown)

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10,70)
$button.Size = New-Object System.Drawing.Size(260,30)
$button.Text = "OK"
$button.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($button)
$form.AcceptButton = $button

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $vmxversion = $dropdown.SelectedItem
    Write-Host "Selected VMX Version: $vmxversion"
    
    # Display all VMs with their hardware versions and host server in a grid view for selection
    $selectedVMs = $allVMs | Out-GridView -Title "Select VMs to Upgrade" -OutputMode Multiple
    
    # Upgrade hardware version of selected VMs and reboot
    foreach ($vm in $selectedVMs) {
        $vmObject = Get-VM -Name $vm.Name
        $spec = New-Object -TypeName VMware.Vim.VirtualMachineConfigSpec
        $spec.ScheduledHardwareUpgradeInfo = New-Object -TypeName VMware.Vim.ScheduledHardwareUpgradeInfo
        $spec.ScheduledHardwareUpgradeInfo.UpgradePolicy = "always"
        $spec.ScheduledHardwareUpgradeInfo.VersionKey = $vmxversion
        $spec.ScheduledHardwareUpgradeInfo.ScheduledHardwareUpgradeStatus = "pending"
        $task = $vmObject.ExtensionData.ReconfigVM_Task($spec)
        
        if ($task -ne $null) {
            Wait-Task -Task $task
            Restart-VMGuest -VM $vmObject -Confirm:$false
            Write-Host "Scheduled hardware upgrade and initiated reboot for VM $($vm.Name)"
        } else {
            Write-Host "Failed to schedule hardware upgrade for VM $($vm.Name)"
        }
    }
}


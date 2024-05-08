
# Load Windows Forms and drawing libraries
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Retrieve all VMs, filter for Windows Servers, and check VMware Tools status
$vmList = Get-VM | Where-Object { $_.Guest.OSFullName -like "*Windows Server*" } | Get-View | Select-Object Name, @{N="VMwareToolsStatus";E={$_.Guest.ToolsStatus}}, @{N="AutoUpgradeOnReboot";E={$_.Config.Tools.ToolsUpgradePolicy -eq "upgradeAtPowerCycle"}}, @{N="ToolsUpgradePending";E={if($_.Guest.ToolsVersionStatus -eq "guestToolsNeedUpgrade"){"Yes"}else{"No"}}} | Sort-Object -Property @{Expression={$_.VMwareToolsStatus -eq "toolsOld"}}, @{Expression="Name"} -Descending

# Display the list in an Out-GridView, allowing for selection
$selectedVMs = $vmList | Out-GridView -Title "Select VMs for Action" -PassThru

# Create the form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Select Action"
$Form.Size = New-Object System.Drawing.Size(300,300) # Adjust size for additional button
$Form.StartPosition = "CenterScreen"

# Define the action buttons
$btnSetAutoTrue = New-Object System.Windows.Forms.Button
$btnReboot = New-Object System.Windows.Forms.Button
$btnUpgradeVMTools = New-Object System.Windows.Forms.Button

# Set properties for the Set Auto to True button
$btnSetAutoTrue.Location = New-Object System.Drawing.Point(10,10)
$btnSetAutoTrue.Size = New-Object System.Drawing.Size(260,30)
$btnSetAutoTrue.Text = "Set Auto to True"
$btnSetAutoTrue.Add_Click({
    PerformAction "SetAutoTrue"
})

# Set properties for the Reboot button
$btnReboot.Location = New-Object System.Drawing.Point(10,50)
$btnReboot.Size = New-Object System.Drawing.Size(260,30)
$btnReboot.Text = "Reboot"
$btnReboot.Add_Click({
    PerformAction "Reboot"
})

# Set properties for the Upgrade VMware Tools button
$btnUpgradeVMTools.Location = New-Object System.Drawing.Point(10,90)
$btnUpgradeVMTools.Size = New-Object System.Drawing.Size(260,30)
$btnUpgradeVMTools.Text = "Upgrade VMware Tools"
$btnUpgradeVMTools.Add_Click({
    PerformAction "UpgradeVMTools"
})

# Add buttons to the form
$Form.Controls.Add($btnSetAutoTrue)
$Form.Controls.Add($btnReboot)
$Form.Controls.Add($btnUpgradeVMTools)

# Progress bar
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(10,130)
$ProgressBar.Size = New-Object System.Drawing.Size(260,20)
$ProgressBar.Minimum = 0
$ProgressBar.Maximum = $selectedVMs.Count
$ProgressBar.Step = 1
$Form.Controls.Add($ProgressBar)

Function PerformAction {
    Param ([string]$action)
    
    $progress = 0
    foreach ($vm in $selectedVMs) {
        $Form.Invoke([Action]{
            $ProgressBar.Value = $progress
        })

        $actualVM = Get-VM -Name $vm.Name
        switch ($action) {
            "SetAutoTrue" {
                $actualVM | Get-View | Foreach-Object {
                    $_.Config.Tools.ToolsUpgradePolicy = "upgradeAtPowerCycle"
                    $_.UpdateViewData()
                }
            }
            "Reboot" {
                Restart-VMGuest -VM $actualVM -Confirm:$false
            }
            "UpgradeVMTools" {
                Update-Tools -VM $actualVM
            }
        }

        $progress++
        Start-Sleep -Seconds 1 # Give some time for the action to be initiated
    }

    $Form.Invoke([Action]{
        $ProgressBar.Value = $selectedVMs.Count
    })
    Start-Sleep -Seconds 2 # Show full progress bar for a moment

    # Disconnect from vCenter and close the form after all actions
   # Disconnect-VIServer -Server $connection -Confirm:$false
    $Form.Close()
}

# Show the form
$Form.Add_Shown({$Form.Activate()})
[void]$Form.ShowDialog()

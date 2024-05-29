# Load required assemblies
Add-Type -AssemblyName PresentationFramework

# Function to show a message box with start/stop options
function Show-ActionPrompt {
    [void][System.Reflection.Assembly]::LoadWithPartialName("PresentationCore")
    [void][System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
    
    $Window = New-Object System.Windows.Window
    $Window.Title = "Action Prompt"
    $Window.SizeToContent = "WidthAndHeight"
    $Window.ResizeMode = "NoResize"
    $Window.WindowStartupLocation = "CenterScreen"
    $Window.Topmost = $true  # Ensure the window stays in the foreground
    
    $StackPanel = New-Object System.Windows.Controls.StackPanel
    $StackPanel.Orientation = "Vertical"
    $StackPanel.HorizontalAlignment = "Center"
    $StackPanel.VerticalAlignment = "Center"
    
    $TextBlock = New-Object System.Windows.Controls.TextBlock
    $TextBlock.Text = "Do you want to start or stop the SSH service on the selected hosts?"
    $TextBlock.Margin = "10"
    $TextBlock.TextAlignment = "Center"
    $StackPanel.Children.Add($TextBlock)
    
    $ButtonStart = New-Object System.Windows.Controls.Button
    $ButtonStart.Content = "Start"
    $ButtonStart.Width = 100
    $ButtonStart.Margin = "5"
    $ButtonStart.Add_Click({
        $global:Action = "start"
        $Window.Close()
    })
    $StackPanel.Children.Add($ButtonStart)
    
    $ButtonStop = New-Object System.Windows.Controls.Button
    $ButtonStop.Content = "Stop"
    $ButtonStop.Width = 100
    $ButtonStop.Margin = "5"
    $ButtonStop.Add_Click({
        $global:Action = "stop"
        $Window.Close()
    })
    $StackPanel.Children.Add($ButtonStop)
    
    $Window.Content = $StackPanel
    $Window.ShowDialog() | Out-Null
}


# Retrieve all ESXi hosts
$esxiHosts = Get-VMHost

# Initialize an array to hold the SSH service status information
$sshServiceStatus = @()

# Check the status of the SSH service on each host
foreach ($esxiHost in $esxiHosts) {
    $sshService = Get-VMHostService -VMHost $esxiHost | Where-Object { $_.Key -eq "TSM-SSH" }
    $status = if ($sshService.Running) { "Running" } else { "Stopped" }
    $sshServiceStatus += [PSCustomObject]@{
        HostName = $esxiHost.Name
        SSHStatus = $status
    }
}

# Display the SSH service status in Out-GridView and allow multiple selections
$selectedHosts = $sshServiceStatus | Out-GridView -Title "SSH Service Status on ESXi Hosts - Select Hosts" -PassThru

if ($selectedHosts) {
    # Show action prompt to start or stop SSH service
    Show-ActionPrompt
    
    if ($global:Action -eq "start" -or $global:Action -eq "stop") {
        # Perform the start or stop action on the selected hosts
        foreach ($selectedHost in $selectedHosts) {
            $esxiHost = Get-VMHost -Name $selectedHost.HostName
            $sshService = Get-VMHostService -VMHost $esxiHost | Where-Object { $_.Key -eq "TSM-SSH" }
            if ($global:Action -eq "start") {
                Start-VMHostService -HostService $sshService -Confirm:$false
            } elseif ($global:Action -eq "stop") {
                Stop-VMHostService -HostService $sshService -Confirm:$false
            }
        }
        
        # Refresh the SSH service status information
        $sshServiceStatus = @()
        foreach ($esxiHost in $esxiHosts) {
            $sshService = Get-VMHostService -VMHost $esxiHost | Where-Object { $_.Key -eq "TSM-SSH" }
            $status = if ($sshService.Running) { "Running" } else { "Stopped" }
            $sshServiceStatus += [PSCustomObject]@{
                HostName = $esxiHost.Name
                SSHStatus = $status
            }
        }
        
        # Display the final SSH service status in Out-GridView
        $sshServiceStatus | Out-GridView -Title "Final SSH Service Status on ESXi Hosts"
    } else {
        Write-Host "Operation cancelled by user."
    }
}


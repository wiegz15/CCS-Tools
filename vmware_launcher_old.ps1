# Import necessary modules and assemblies for UI and VMware management
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# Check and install ImportExcel module if necessary
$importExcelModule = Get-Module -ListAvailable -Name ImportExcel
if (-not $importExcelModule) {
    $response = [System.Windows.Forms.MessageBox]::Show("The ImportExcel module is not installed. Do you want to install it now?", "Module Not Found", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($response -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force
            Import-Module ImportExcel
            [System.Windows.Forms.MessageBox]::Show("ImportExcel module installed successfully.", "Installation Successful")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to install ImportExcel module. Error: $($_.Exception.Message)", "Installation Failed")
            return
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("ImportExcel module is required to continue. Exiting script...", "Error")
        return
    }
} else {
    Import-Module ImportExcel
}

# Check and install VMware PowerCLI if necessary
$powerCLIModule = Get-Module -Name VMware.PowerCLI -ListAvailable
$latestPowerCLI = Find-Module -Name VMware.PowerCLI

if (-not $powerCLIModule) {
    $installResponse = [System.Windows.Forms.MessageBox]::Show("VMware PowerCLI is not installed. Do you want to install it now?", "PowerCLI Not Found", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($installResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Install-Module -Name VMware.PowerCLI -Scope CurrentUser -Force -AllowClobber
            Import-Module VMware.PowerCLI
            [System.Windows.Forms.MessageBox]::Show("VMware PowerCLI installed successfully.", "Installation Successful")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to install VMware PowerCLI. Error: $($_.Exception.Message)", "Installation Failed")
            return
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("VMware PowerCLI is required to continue. Exiting script...", "Error")
        return
    }
} elseif ($powerCLIModule.Version -lt $latestPowerCLI.Version) {
    $updateResponse = [System.Windows.Forms.MessageBox]::Show("A newer version of VMware PowerCLI is available. Do you want to update now?", "Update Available", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($updateResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Update-Module -Name VMware.PowerCLI -Force
            [System.Windows.Forms.MessageBox]::Show("VMware PowerCLI updated successfully.", "Update Successful")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to update VMware PowerCLI. Error: $($_.Exception.Message)", "Update Failed")
            return
        }
    }
} else {
    Import-Module VMware.PowerCLI
}

# Function to bring the PowerShell console to the foreground
function Bring-ConsoleToFront {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 9)  # 9 is Restore
    [Console.Window]::SetForegroundWindow($consolePtr)
}

# Define paths based on the script location
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
$binDir = Join-Path -Path $scriptDirectory -ChildPath "bin"
$reportsDir = Join-Path -Path $scriptDirectory -ChildPath "Reports"
$vmwareDir = Join-Path -Path $scriptDirectory -ChildPath "VMware"

# Ensure the bin directory exists
if (-not (Test-Path $binDir)) {
    New-Item -Path $binDir -ItemType Directory
}

# Ensure the Reports directory exists
if (-not (Test-Path $reportsDir)) {
    New-Item -Path $reportsDir -ItemType Directory
}

# Ensure the VMware directory exists
if (-not (Test-Path $vmwareDir)) {
    New-Item -Path $vmwareDir -ItemType Directory
    [System.Windows.Forms.MessageBox]::Show("VMware directory not found. A new VMware directory has been created.", "Directory Created")
}

# Define the path for the tools.ini file within the bin directory
$toolsIniPath = Join-Path -Path $binDir -ChildPath "tools.ini"

# Check if the tools.ini file exists
if (Test-Path $toolsIniPath) {
    $iniContent = Get-Content $toolsIniPath
    $serverName = ($iniContent | Where-Object { $_ -match "vcenter=" }).Split('=')[1]
    if (-not $serverName) {
        $inputServerName = [System.Windows.Forms.MessageBox]::Show("vCenter server name not found. Enter vCenter server name:", "Enter Server Name", [System.Windows.Forms.MessageBoxButtons]::OKCancel)
        if ($inputServerName -eq [System.Windows.Forms.DialogResult]::OK) {
            $serverName = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the vCenter server name", "Server Name Input", "")
            if (-not $serverName) {
                [System.Windows.Forms.MessageBox]::Show("No server name entered. Exiting script...", "Error")
                return
            } else {
                "vcenter=$serverName" | Set-Content $toolsIniPath
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Operation cancelled by the user.", "Cancelled")
            return
        }
    }
} else {
    $createNew = [System.Windows.Forms.MessageBox]::Show("Vcenter Server Address file not found. Do you want to create it now?", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($createNew -eq [System.Windows.Forms.DialogResult]::Yes) {
        $serverName = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the vCenter server name", "Server Name Input", "")
        if (-not $serverName) {
            [System.Windows.Forms.MessageBox]::Show("No server name entered. Exiting script...", "Error")
            return
        } else {
            "vcenter=$serverName" | Set-Content $toolsIniPath
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Cannot proceed without a Vcenter Address file. Exiting script...", "Error")
        return
    }
}

# Check server availability on port 443
$client = New-Object System.Net.Sockets.TcpClient
try {
    $client.Connect($serverName, 443)
    $client.Close()
} catch {
    [System.Windows.Forms.MessageBox]::Show("$serverName is not online on port 443.", "Connection Check Failed")
    return
}

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "Output.xlsx"

# Delete the Output.xlsx file if it exists at the start of the script
if (Test-Path $excelPath) {
    Remove-Item $excelPath -Force
}

# Define variable to track connection status
$connected = $false

# Setup the main window for the launcher
$window = New-Object System.Windows.Window
$window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
$window.Left = 0
$window.Top = 0
$window.Title = "VMware Tools"
$window.Width = 400
$window.Height = 860
$window.Topmost = $true

# Add event handler for when the window is closing
$window.Add_Closing({
    if ($connected) {
        Disconnect-VIServer -Server $serverName -Confirm:$false
        [System.Windows.Forms.MessageBox]::Show("Disconnected from $serverName", "Disconnected")
    }
})

# ScrollViewer for the StackPanel
$scrollViewer = New-Object System.Windows.Controls.ScrollViewer
$scrollViewer.VerticalScrollBarVisibility = "Auto"
$scrollViewer.Margin = 10
$scrollViewer.Height = $window.Height
$mainStackPanel = New-Object System.Windows.Controls.StackPanel
$scrollViewer.Content = $mainStackPanel

# Label to show connected vCenter address
$vCenterAddressLabel = New-Object System.Windows.Controls.TextBlock
$vCenterAddressLabel.Text = ""
$vCenterAddressLabel.Margin = "10,5,10,5"
$vCenterAddressLabel.Foreground = "Black"
$mainStackPanel.Children.Add($vCenterAddressLabel)

# Connect button to connect to vCenter Server
$connectButton = New-Object System.Windows.Controls.Button
$connectButton.Content = "Connect to vCenter"
$connectButton.Margin = 10
# Connect button functionality
$connectButton.Add_Click({
    try {
        $credentials = Get-Credential
        if ($null -eq $credentials) {
            [System.Windows.Forms.MessageBox]::Show("No credentials entered. Exiting script...", "Error")
            return
        }
        Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
        Connect-VIServer -Server $serverName -Credential $credentials -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("Successfully connected to $serverName", "Connection Successful")
        $connected = $true
        $vCenterAddressLabel.Text = "Connected to: $serverName"
        $executeButton.IsEnabled = $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to $serverName. Error: $($_.Exception.Message)", "Connection Failed")
        return
    }
})

$mainStackPanel.Children.Add($connectButton)

# "Select All" Checkbox
$selectAllCheckbox = New-Object System.Windows.Controls.CheckBox
$selectAllCheckbox.Content = "Select All"
$selectAllCheckbox.Margin = 10
$selectAllCheckbox.Add_Checked({
    $mainStackPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckbox } | ForEach-Object {
        $_.IsChecked = $true
    }
})
$selectAllCheckbox.Add_Unchecked({
    $mainStackPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckbox } | ForEach-Object {
        $_.IsChecked = $false
    }
})
$mainStackPanel.Children.Add($selectAllCheckbox)

# Adding CheckBoxes for scripts
$scriptFiles = Get-ChildItem -Path $vmwareDir -Filter *.ps1 | Sort-Object Name
foreach ($scriptFile in $scriptFiles) {
    $checkBox = New-Object System.Windows.Controls.CheckBox
    $checkBox.Content = $scriptFile.Name
    $checkBox.Margin = 5
    $mainStackPanel.Children.Add($checkBox)
}

# Execute button to run selected scripts
$executeButton = New-Object System.Windows.Controls.Button
$executeButton.Content = "Execute"
$executeButton.Margin = 10
$executeButton.IsEnabled = $false
$executeButton.Add_Click({
    # Collect only script checkboxes, ignoring the "Select All" checkbox
    $selectedScripts = $mainStackPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked -and $_.Content -ne 'Select All' }
    foreach ($selectedScript in $selectedScripts) {
        $scriptPath = Join-Path $vmwareDir $selectedScript.Content
        . $scriptPath
    }
    [System.Windows.Forms.MessageBox]::Show("Scripts execution completed", "Execution Complete")

    # Prompt user to open the Output.xlsx file
    $openResponse = [System.Windows.Forms.MessageBox]::Show("Do you want to open the Output.xlsx file now?", "Open File?", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($openResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
        $excelPath = Join-Path -Path $reportsDir -ChildPath "Output.xlsx"
        if (Test-Path $excelPath) {
            Start-Process $excelPath  # Opens the file with the default application associated with .xlsx files
        } else {
            [System.Windows.Forms.MessageBox]::Show("Output.xlsx file not found.", "File Not Found")
        }
    }
})
$mainStackPanel.Children.Add($executeButton)

# Open Report button
$openReportButton = New-Object System.Windows.Controls.Button
$openReportButton.Content = "Open Report"
$openReportButton.Margin = 10
$openReportButton.Add_Click({
    if (Test-Path $excelPath) {
        Start-Process $excelPath  # Opens the file with the default application associated with .xlsx files
    } else {
        [System.Windows.Forms.MessageBox]::Show("Output.xlsx file not found.", "File Not Found")
    }
})
$mainStackPanel.Children.Add($openReportButton)

# Reset button to delete the Output.xlsx file
$resetButton = New-Object System.Windows.Controls.Button
$resetButton.Content = "Reset"
$resetButton.Margin = 10
$resetButton.Add_Click({
    if (Test-Path $excelPath) {
        Remove-Item $excelPath -Force
    } else {
        [System.Windows.Forms.MessageBox]::Show("No Output.xlsx file found.", "File Not Found")
    }
})
$mainStackPanel.Children.Add($resetButton)

# Exit button to close the launcher
$exitButton = New-Object System.Windows.Controls.Button
$exitButton.Content = "Exit"
$exitButton.Margin = 10
$exitButton.Add_Click({
    if ($connected) {
        Disconnect-VIServer -Server $serverName -Confirm:$false
        [System.Windows.Forms.MessageBox]::Show("Disconnected from $serverName", "Disconnected")
    }
    $window.Close()
})
$mainStackPanel.Children.Add($exitButton)

# Set the window content and show the window
$window.Content = $scrollViewer
$window.ShowDialog()

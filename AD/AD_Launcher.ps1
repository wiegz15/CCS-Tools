# Import necessary modules and assemblies for UI and VMware management
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

function Show-ProgressForm {
    param($title = "Processing", $message = "Please wait...")

    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = $title
    $progressForm.Size = New-Object System.Drawing.Size(300, 100)
    $progressForm.StartPosition = "CenterScreen"
    $progressForm.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $message
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(280, 20)
    $progressForm.Controls.Add($label)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 40)
    $progressBar.Size = New-Object System.Drawing.Size(260, 20)
    $progressBar.Style = "Continuous"
    $progressForm.Controls.Add($progressBar)

    $progressForm.Show()

    # Return objects
    return @{ Form = $progressForm; ProgressBar = $progressBar; Label = $label }
}

function Update-ProgressBar {
    param($progressData, $text, $percent)
    $progressData.Label.Text = $text
    $progressData.ProgressBar.Value = $percent
    $progressData.Form.Refresh()
}

# Function to bring the PowerShell console to the foreground
function Bring-ConsoleToFront {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 9)  # 9 is Restore
    [Console.Window]::SetForegroundWindow($consolePtr)
}

# Define paths based on the script location
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
$parentDirectory = Split-Path -Parent $scriptDirectory
$reportsDir = Join-Path -Path $parentDirectory -ChildPath "Reports"
$adScriptsDir = Join-Path -Path $scriptDirectory -ChildPath "Scripts"
$Otherscripts = Join-Path -Path $adScriptsDir -ChildPath "Other"

# Ensure the Reports directory exists
if (-not (Test-Path $reportsDir)) {
    New-Item -Path $reportsDir -ItemType Directory
}

# # Ensure the VMware directory exists
# if (-not (Test-Path $extrascripts)) {
#     New-Item -Path $extrascripts -ItemType Directory
#     [System.Windows.Forms.MessageBox]::Show("Extra directory not found. A new Extra directory has been created.", "Directory Created")
# }

# Ensure the EXTRA Scripts directory exists
if (-not (Test-Path $adScriptsDir)) {
    New-Item -Path $adScriptsDir -ItemType Directory
    [System.Windows.Forms.MessageBox]::Show("AD Scripts directory not found. A new AD Scripts directory has been created.", "Directory Created")
}

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"

# Delete the Output.xlsx file if it exists at the start of the script
if (Test-Path $excelPath) {
    Remove-Item $excelPath -Force
}

# Define the main window
$window = New-Object System.Windows.Window
$window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
$window.Left = 0
$window.Top = 0
$window.Title = "AD Tools"
$window.Width = 400
$window.Height = 860
$window.Topmost = $true

# Main container StackPanel
$mainContainer = New-Object System.Windows.Controls.StackPanel

function UpdateUIElements {
    # Enable or disable execute buttons
    $executeButton1.IsEnabled = $global:connected
    $executeButton2.IsEnabled = $global:connected

    # Enable or disable checkboxes in both tabs
    $mainStackPanel1.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] } | ForEach-Object {
        $_.IsEnabled = $global:connected
    }
    $mainStackPanel2.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] } | ForEach-Object {
        $_.IsEnabled = $global:connected
    }
}

# TabControl for different sets of tools
$tabControl = New-Object System.Windows.Controls.TabControl
$tabControl.Margin = 10

# Tab 1: AD Health
$tabItem1 = New-Object System.Windows.Controls.TabItem
$tabItem1.Header = "AD Health"
$scrollViewer1 = New-Object System.Windows.Controls.ScrollViewer
$scrollViewer1.VerticalScrollBarVisibility = "Auto"
$scrollViewer1.HorizontalAlignment = "Stretch"
$scrollViewer1.VerticalAlignment = "Stretch"
$scrollViewer1.Margin = 10
$mainStackPanel1 = New-Object System.Windows.Controls.StackPanel
$scrollViewer1.Content = $mainStackPanel1
$tabItem1.Content = $scrollViewer1

# "Select All" Checkbox
$selectAllCheckbox = New-Object System.Windows.Controls.CheckBox
$selectAllCheckbox.Content = "Select All"
$selectAllCheckbox.Margin = 10
$selectAllCheckbox.Add_Checked({
    $mainStackPanel1.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckbox } | ForEach-Object {
        $_.IsChecked = $true
    }
})
$selectAllCheckbox.Add_Unchecked({
    $mainStackPanel1.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckbox } | ForEach-Object {
        $_.IsChecked = $false
    }
})
$mainStackPanel1.Children.Add($selectAllCheckbox)

# Load scripts for Tab 1
$scriptFiles1 = Get-ChildItem -Path $adScriptsDir -Filter *.ps1 | Sort-Object Name
foreach ($scriptFile in $scriptFiles1) {
    $checkBox = New-Object System.Windows.Controls.CheckBox
    $checkBox.Content = $scriptFile.Name
    $checkBox.Margin = 5
    $mainStackPanel1.Children.Add($checkBox)
}

$executeButton1 = New-Object System.Windows.Controls.Button
$executeButton1.Content = "Execute"
$executeButton1.Margin = 5
$executeButton1.Add_Click({
    # Collect only script checkboxes, ignoring the "Select All" checkbox
    $selectedScripts = $mainStackPanel1.Children | Where-Object {
        $_ -is [System.Windows.Controls.CheckBox] -and
        $_.IsChecked -and
        $_.Content -ne 'Select All'
    }
    
    if ($selectedScripts.Count -gt 0) {
        $progressData = Show-ProgressForm -title "Executing Scripts" -message "Starting script executions..."
        $totalScripts = $selectedScripts.Count
        $currentScriptIndex = 0

        foreach ($selectedScript in $selectedScripts) {
            $scriptPath = Join-Path -Path $adScriptsDir -ChildPath $selectedScript.Content
            Update-ProgressBar -progressData $progressData -text "Running $($selectedScript.Content)..." -percent (($currentScriptIndex / $totalScripts) * 100)
            . $scriptPath
            $currentScriptIndex++
        }

        Update-ProgressBar -progressData $progressData -text "Scripts execution completed" -percent 100
        Start-Sleep -Seconds 2
        $progressData.Form.Close()
        [System.Windows.Forms.MessageBox]::Show("Scripts execution completed", "Execution Complete")
        
    }

    # Prompt user to open the Output.xlsx file
    $openResponse = [System.Windows.Forms.MessageBox]::Show("Do you want to open the AD_Output.xlsx file now?", "Open File?", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($openResponse -eq [System.Windows.Forms.DialogResult]::Yes) {
        $excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
        if (Test-Path $excelPath) {
            Start-Process $excelPath  # Opens the file with the default application associated with .xlsx files
        } else {
            [System.Windows.Forms.MessageBox]::Show("AD_Output.xlsx file not found.", "File Not Found")
        }
    }
})

$mainStackPanel1.Children.Add($executeButton1)

# Open Report button
$openReportButton = New-Object System.Windows.Controls.Button
$openReportButton.Content = "Open Report"
$openReportButton.Margin = 5
$openReportButton.Add_Click({
    if (Test-Path $excelPath) {
        Start-Process $excelPath  # Opens the file with the default application associated with .xlsx files
    } else {
        [System.Windows.Forms.MessageBox]::Show("AD_Output.xlsx file not found.", "File Not Found")
    }
})
$mainStackPanel1.Children.Add($openReportButton)

# Reset button to delete the Output.xlsx file
$resetButton = New-Object System.Windows.Controls.Button
$resetButton.Content = "Reset"
$resetButton.Margin = 5
$resetButton.Add_Click({
    if (Test-Path $excelPath) {
        Remove-Item $excelPath -Force
    } else {
        [System.Windows.Forms.MessageBox]::Show("AD_Output.xlsx file found and cleared.", "File Not Found")
    }
})
$mainStackPanel1.Children.Add($resetButton)

# Tab 2: Other Reports
$tabItem2 = New-Object System.Windows.Controls.TabItem
$tabItem2.Header = "Other"
$scrollViewer2 = New-Object System.Windows.Controls.ScrollViewer
$scrollViewer2.VerticalScrollBarVisibility = "Auto"
$scrollViewer2.HorizontalAlignment = "Stretch"
$scrollViewer2.VerticalAlignment = "Stretch"
$scrollViewer2.Margin = 10
$mainStackPanel2 = New-Object System.Windows.Controls.StackPanel
$scrollViewer2.Content = $mainStackPanel2
$tabItem2.Content = $scrollViewer2

# "Select All" Checkbox
$selectAllCheckbox = New-Object System.Windows.Controls.CheckBox
$selectAllCheckbox.Content = "Select All"
$selectAllCheckbox.Margin = 10
$selectAllCheckbox.Add_Checked({
    $mainStackPanel2.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckbox } | ForEach-Object {
        $_.IsChecked = $true
    }
})
$selectAllCheckbox.Add_Unchecked({
    $mainStackPanel2.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckbox } | ForEach-Object {
        $_.IsChecked = $false
    }
})
$mainStackPanel2.Children.Add($selectAllCheckbox)

# Load scripts for Tab 2
$scriptFiles2 = Get-ChildItem -Path $Otherscripts -Filter *.ps1 | Sort-Object Name
foreach ($scriptFile in $scriptFiles2) {
    $checkBox = New-Object System.Windows.Controls.CheckBox
    $checkBox.Content = $scriptFile.Name
    $checkBox.Margin = 5
    $mainStackPanel2.Children.Add($checkBox)
}

$executeButton2 = New-Object System.Windows.Controls.Button
$executeButton2.Content = "Execute"
$executeButton2.Margin = 5
$executeButton2.Add_Click({
    # Collect only script checkboxes, ignoring the "Select All" checkbox
    $selectedScripts = $mainStackPanel2.Children | Where-Object {
        $_ -is [System.Windows.Controls.CheckBox] -and
        $_.IsChecked -and
        $_.Content -ne 'Select All'
    }
    
    if ($selectedScripts.Count -gt 0) {
        $progressData = Show-ProgressForm -title "Executing Scripts" -message "Starting script executions..."
        $totalScripts = $selectedScripts.Count
        $currentScriptIndex = 0

        foreach ($selectedScript in $selectedScripts) {
            $scriptPath = Join-Path -Path $Otherscripts -ChildPath $selectedScript.Content
            Update-ProgressBar -progressData $progressData -text "Running $($selectedScript.Content)..." -percent (($currentScriptIndex / $totalScripts) * 100)
            . $scriptPath
            $currentScriptIndex++
        }

        Update-ProgressBar -progressData $progressData -text "Scripts execution completed" -percent 100
        Start-Sleep -Seconds 2
        $progressData.Form.Close()        
    }
})

$mainStackPanel2.Children.Add($executeButton2)

# Adding tabs to TabControl
$tabControl.Items.Add($tabItem1)
$tabControl.Items.Add($tabItem2) #Enable to Enable Tab 2

# Add the TabControl to the main container
$mainContainer.Children.Add($tabControl)

# Exit button to close the launcher
$exitButton = New-Object System.Windows.Controls.Button
$exitButton.Content = "Exit"
$exitButton.Margin = 5
$exitButton.Add_Click({
    if ($connected) {
        Disconnect-VIServer -Server $serverName -Confirm:$false
        [System.Windows.Forms.MessageBox]::Show("Disconnected from $serverName", "Disconnected")
    }
    $window.Close()
})
$mainContainer.Children.Add($exitButton)

# Set the main container as the window content
$window.Content = $mainContainer
$window.ShowDialog()
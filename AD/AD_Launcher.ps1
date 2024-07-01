# Import necessary modules and assemblies for UI and management
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
$Otherscripts = Join-Path -Path $adScriptsDir -ChildPath "WorkinProgress"
$ExchangeSCriptsDir = Join-Path -Path $adScriptsDir -ChildPath "ExchangeScripts"

# Ensure the Reports directory exists
if (-not (Test-Path $reportsDir)) {
    New-Item -Path $reportsDir -ItemType Directory
}

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
# Updated Execute button click event for Tab 1: AD Health
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

        $scriptCommands = "`$reportsDir = `"$reportsDir`";"

        foreach ($selectedScript in $selectedScripts) {
            $scriptPath = Join-Path -Path $adScriptsDir -ChildPath $selectedScript.Content
            $scriptCommands += "& `"$scriptPath`"; "
            Update-ProgressBar -progressData $progressData -text "Adding $($selectedScript.Content) to batch..." -percent (($currentScriptIndex / $totalScripts) * 100)
            $currentScriptIndex++
        }

        $tempScriptPath = Join-Path -Path $env:TEMP -ChildPath "BatchScript.ps1"
        Set-Content -Path $tempScriptPath -Value $scriptCommands

        Update-ProgressBar -progressData $progressData -text "Running all selected scripts in a new window..." -percent 100
        Start-Sleep -Seconds 1
        $progressData.Form.Close()

        # Open a new PowerShell window to run the combined script
        Start-Process powershell.exe -ArgumentList "-NoExit -File `"$tempScriptPath`""
    }
})

$mainStackPanel1.Children.Add($executeButton1)

# Open Report button
$openReportButton = New-Object System.Windows.Controls.Button
$openReportButton.Content = "Open Report Folder"
$openReportButton.Margin = 5
$openReportButton.Add_Click({
    if (Test-Path $reportsDir) {
        Start-Process explorer.exe -ArgumentList "`"$reportsDir`""  # Opens the folder containing the report file
    } else {
        [System.Windows.Forms.MessageBox]::Show("Reports directory not found.", "Directory Not Found")
    }
})
$mainStackPanel1.Children.Add($openReportButton)

# Tab 2: Other Reports
$tabItem2 = New-Object System.Windows.Controls.TabItem
$tabItem2.Header = "WIP"
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

        $scriptCommands = "`$reportsDir = `"$reportsDir`";"

        foreach ($selectedScript in $selectedScripts) {
            $scriptPath = Join-Path -Path $Otherscripts -ChildPath $selectedScript.Content
            $scriptCommands += "& `"$scriptPath`"; "
            Update-ProgressBar -progressData $progressData -text "Adding $($selectedScript.Content) to batch..." -percent (($currentScriptIndex / $totalScripts) * 100)
            $currentScriptIndex++
        }

        $tempScriptPath = Join-Path -Path $env:TEMP -ChildPath "BatchScript.ps1"
        Set-Content -Path $tempScriptPath -Value $scriptCommands

        Update-ProgressBar -progressData $progressData -text "Running all selected scripts in a new window..." -percent 100
        Start-Sleep -Seconds 1
        $progressData.Form.Close()

        # Open a new PowerShell window to run the combined script
        Start-Process powershell.exe -ArgumentList "-NoExit -File `"$tempScriptPath`""
    }
})

$mainStackPanel2.Children.Add($executeButton2)

$openReportButton2 = New-Object System.Windows.Controls.Button
$openReportButton2.Content = "Open Report Folder"
$openReportButton2.Margin = 5
$openReportButton2.Add_Click({
    if (Test-Path $reportsDir) {
        Start-Process explorer.exe -ArgumentList "`"$reportsDir`""  # Opens the folder containing the report file
    } else {
        [System.Windows.Forms.MessageBox]::Show("Reports directory not found.", "Directory Not Found")
    }
})
$mainStackPanel2.Children.Add($openReportButton2)

# Tab 3: Other Reports
$tabItem3 = New-Object System.Windows.Controls.TabItem
$tabItem3.Header = "Exchange"
$scrollViewer3 = New-Object System.Windows.Controls.ScrollViewer
$scrollViewer3.VerticalScrollBarVisibility = "Auto"
$scrollViewer3.HorizontalAlignment = "Stretch"
$scrollViewer3.VerticalAlignment = "Stretch"
$scrollViewer3.Margin = 10
$mainStackPanel3 = New-Object System.Windows.Controls.StackPanel
$scrollViewer3.Content = $mainStackPanel3
$tabItem3.Content = $scrollViewer3

# "Select All" Checkbox
$selectAllCheckbox = New-Object System.Windows.Controls.CheckBox
$selectAllCheckbox.Content = "Select All"
$selectAllCheckbox.Margin = 10
$selectAllCheckbox.Add_Checked({
    $mainStackPanel3.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckbox } | ForEach-Object {
        $_.IsChecked = $true
    }
})
$selectAllCheckbox.Add_Unchecked({
    $mainStackPanel3.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckbox } | ForEach-Object {
        $_.IsChecked = $false
    }
})
$mainStackPanel3.Children.Add($selectAllCheckbox)

# Load scripts for Tab 3
$scriptFiles3 = Get-ChildItem -Path $ExchangeSCriptsDir -Filter *.ps1 | Sort-Object Name
foreach ($scriptFile in $scriptFiles3) {
    $checkBox = New-Object System.Windows.Controls.CheckBox
    $checkBox.Content = $scriptFile.Name
    $checkBox.Margin = 5
    $mainStackPanel3.Children.Add($checkBox)
}

$executeButton3 = New-Object System.Windows.Controls.Button
$executeButton3.Content = "Execute"
$executeButton3.Margin = 5
$executeButton3.Add_Click({
    # Collect only script checkboxes, ignoring the "Select All" checkbox
    $selectedScripts = $mainStackPanel3.Children | Where-Object {
        $_ -is [System.Windows.Controls.CheckBox] -and
        $_.IsChecked -and
        $_.Content -ne 'Select All'
    }
    
    if ($selectedScripts.Count -gt 0) {
        $progressData = Show-ProgressForm -title "Executing Scripts" -message "Starting script executions..."
        $totalScripts = $selectedScripts.Count
        $currentScriptIndex = 0

        $scriptCommands = "`$reportsDir = `"$reportsDir`";"

        foreach ($selectedScript in $selectedScripts) {
            $scriptPath = Join-Path -Path $ExchangeSCriptsDir -ChildPath $selectedScript.Content
            $scriptCommands += "& `"$scriptPath`"; "
            Update-ProgressBar -progressData $progressData -text "Adding $($selectedScript.Content) to batch..." -percent (($currentScriptIndex / $totalScripts) * 100)
            $currentScriptIndex++
        }

        $tempScriptPath = Join-Path -Path $env:TEMP -ChildPath "BatchScript.ps1"
        Set-Content -Path $tempScriptPath -Value $scriptCommands

        Update-ProgressBar -progressData $progressData -text "Running all selected scripts in a new window..." -percent 100
        Start-Sleep -Seconds 1
        $progressData.Form.Close()

        # Open a new PowerShell window to run the combined script
        Start-Process powershell.exe -ArgumentList "-NoExit -File `"$tempScriptPath`""
    }
})

$mainStackPanel3.Children.Add($executeButton3)

$openReportButton3 = New-Object System.Windows.Controls.Button
$openReportButton3.Content = "Open Report Folder"
$openReportButton3.Margin = 5
$openReportButton3.Add_Click({
    if (Test-Path $reportsDir) {
        Start-Process explorer.exe -ArgumentList "`"$reportsDir`""  # Opens the folder containing the report file
    } else {
        [System.Windows.Forms.MessageBox]::Show("Reports directory not found.", "Directory Not Found")
    }
})
$mainStackPanel3.Children.Add($openReportButton3)



# Adding tabs to TabControl
$tabControl.Items.Add($tabItem1)
$tabControl.Items.Add($tabItem2)
$tabControl.Items.Add($tabItem3)

# Add the TabControl to the main container
$mainContainer.Children.Add($tabControl)

# ClearReports button to delete the Output.xlsx file
$ClearReports = New-Object System.Windows.Controls.Button
$ClearReports.Content = "Clear Reports Folder"
$ClearReports.Margin = 5
$ClearReports.Add_Click({
    ShowFileSelectionPopup
})
$mainContainer.Children.Add($ClearReports)

# Function to show the file selection popup
function ShowFileSelectionPopup {
    # Create the popup window
    $popupWindow = New-Object System.Windows.Window
    $popupWindow.Title = "Select Files to Delete"
    $popupWindow.Width = 400
    $popupWindow.Height = 300
    $popupWindow.WindowStartupLocation = 'CenterScreen'

    # Create a StackPanel to hold the controls in the popup window
    $popupStackPanel = New-Object System.Windows.Controls.StackPanel
    $popupStackPanel.Orientation = "Vertical"
    $popupStackPanel.Margin = 10

    # Create a ListBox to display the files
    $listBox = New-Object System.Windows.Controls.ListBox
    $listBox.SelectionMode = "Extended"
    $popupStackPanel.Children.Add($listBox)

    # Create a button to delete the selected files
    $deleteButton = New-Object System.Windows.Controls.Button
    $deleteButton.Content = "Delete Selected Files"
    $deleteButton.Margin = 5
    $deleteButton.Add_Click({
        $selectedItems = $listBox.SelectedItems
        if ($selectedItems.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No files selected.", "Warning")
        } else {
            foreach ($item in $selectedItems) {
                Remove-Item $item -Force
            }
            [System.Windows.MessageBox]::Show("Selected files deleted.", "Success")
            RefreshFileList $listBox
        }
    })
    $popupStackPanel.Children.Add($deleteButton)

    # Create a button to refresh the file list
    $refreshButton = New-Object System.Windows.Controls.Button
    $refreshButton.Content = "Refresh File List"
    $refreshButton.Margin = 5
    $refreshButton.Add_Click({
        RefreshFileList $listBox
    })
    $popupStackPanel.Children.Add($refreshButton)

    # Add the StackPanel to the popup window
    $popupWindow.Content = $popupStackPanel

    # Function to refresh the file list
    function RefreshFileList {
        param ($listBox)
        $listBox.Items.Clear()
        $files = Get-ChildItem -Path $reportsDir
        foreach ($file in $files) {
            $listBox.Items.Add($file.Name)
        }
    }

    # Initial load of file list
    RefreshFileList $listBox

    # Show the popup window
    $popupWindow.ShowDialog() | Out-Null
}

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
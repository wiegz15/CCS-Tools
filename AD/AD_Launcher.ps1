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

# Ensure the AD Scripts directory exists
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

# Define a function to refresh the scripts loaded in the UI
function Refresh-Scripts {
    # Clear the existing children (checkboxes) in both tab panels
    $mainStackPanel1.Children.Clear()
    $mainStackPanel2.Children.Clear()

    # Create new "Select All" checkbox for Tab 1
    $selectAllCheckboxTab1 = New-Object System.Windows.Controls.CheckBox
    $selectAllCheckboxTab1.Content = "Select All"
    $selectAllCheckboxTab1.Margin = 10
    $selectAllCheckboxTab1.Add_Checked({
        $mainStackPanel1.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckboxTab1 } | ForEach-Object {
            $_.IsChecked = $true
        }
    })
    $selectAllCheckboxTab1.Add_Unchecked({
        $mainStackPanel1.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckboxTab1 } | ForEach-Object {
            $_.IsChecked = $false
        }
    })
    $mainStackPanel1.Children.Add($selectAllCheckboxTab1)

    # Load scripts for Tab 1 (AD Health)
    $scriptFiles1 = Get-ChildItem -Path $adScriptsDir -Filter *.ps1 | Sort-Object Name
    foreach ($scriptFile in $scriptFiles1) {
        $checkBox = New-Object System.Windows.Controls.CheckBox
        $checkBox.Content = $scriptFile.Name
        $checkBox.Margin = 5
        $mainStackPanel1.Children.Add($checkBox)
    }

    # Re-add the Execute button for Tab 1 after reloading scripts
    $mainStackPanel1.Children.Add($executeButton1)

    # Create new "Select All" checkbox for Tab 2
    $selectAllCheckboxTab2 = New-Object System.Windows.Controls.CheckBox
    $selectAllCheckboxTab2.Content = "Select All"
    $selectAllCheckboxTab2.Margin = 10
    $selectAllCheckboxTab2.Add_Checked({
        $mainStackPanel2.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckboxTab2 } | ForEach-Object {
            $_.IsChecked = $true
        }
    })
    $selectAllCheckboxTab2.Add_Unchecked({
        $mainStackPanel2.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_ -ne $selectAllCheckboxTab2 } | ForEach-Object {
            $_.IsChecked = $false
        }
    })
    $mainStackPanel2.Children.Add($selectAllCheckboxTab2)

    # Load scripts for Tab 2 (Other)
    $scriptFiles2 = Get-ChildItem -Path $Otherscripts -Filter *.ps1 | Sort-Object Name
    foreach ($scriptFile in $scriptFiles2) {
        $checkBox = New-Object System.Windows.Controls.CheckBox
        $checkBox.Content = $scriptFile.Name
        $checkBox.Margin = 5
        $mainStackPanel2.Children.Add($checkBox)
    }

    # Re-add the Execute button for Tab 2 after reloading scripts
    $mainStackPanel2.Children.Add($executeButton2)
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

# Scroll viewer for each tab
$scrollViewer1 = New-Object System.Windows.Controls.ScrollViewer
$scrollViewer1.VerticalScrollBarVisibility = "Auto"
$scrollViewer1.HorizontalAlignment = "Stretch"
$scrollViewer1.VerticalAlignment = "Stretch"
$scrollViewer1.Margin = 10
$mainStackPanel1 = New-Object System.Windows.Controls.StackPanel
$scrollViewer1.Content = $mainStackPanel1

$scrollViewer2 = New-Object System.Windows.Controls.ScrollViewer
$scrollViewer2.VerticalScrollBarVisibility = "Auto"
$scrollViewer2.HorizontalAlignment = "Stretch"
$scrollViewer2.VerticalAlignment = "Stretch"
$scrollViewer2.Margin = 10
$mainStackPanel2 = New-Object System.Windows.Controls.StackPanel
$scrollViewer2.Content = $mainStackPanel2

# TabControl for different sets of tools
$tabControl = New-Object System.Windows.Controls.TabControl
$tabControl.Margin = 10

# Tab 1: AD Health
$tabItem1 = New-Object System.Windows.Controls.TabItem
$tabItem1.Header = "AD Health"
$tabItem1.Content = $scrollViewer1
$tabControl.Items.Add($tabItem1)

# Tab 2: Other Reports
$tabItem2 = New-Object System.Windows.Controls.TabItem
$tabItem2.Header = "Other"
$tabItem2.Content = $scrollViewer2
$tabControl.Items.Add($tabItem2)

# Add the TabControl to the main container
$mainContainer.Children.Add($tabControl)

# Adding a refresh button
$refreshButton = New-Object System.Windows.Controls.Button
$refreshButton.Content = "Refresh Scripts"
$refreshButton.Margin = 5
$refreshButton.Add_Click({
    Refresh-Scripts
    [System.Windows.Forms.MessageBox]::Show("Scripts reloaded successfully.", "Refreshed")
})

# Add the refresh button to the main container
$mainContainer.Children.Add($refreshButton)

# Execute button for Tab 1
$executeButton1 = New-Object System.Windows.Controls.Button
$executeButton1.Content = "Execute"
$executeButton1.Margin = 5
$executeButton1.Add_Click({
    $selectedScripts = $mainStackPanel1.Children | Where-Object {
        $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked -and $_.Content -ne 'Select All'
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
})

$mainStackPanel1.Children.Add($executeButton1)

# Execute button for Tab 2
$executeButton2 = New-Object System.Windows.Controls.Button
$executeButton2.Content = "Execute"
$executeButton2.Margin = 5
$executeButton2.Add_Click({
    $selectedScripts = $mainStackPanel2.Children | Where-Object {
        $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked -and $_.Content -ne 'Select All'
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

# Exit button to close the launcher
$exitButton = New-Object System.Windows.Controls.Button
$exitButton.Content = "Exit"
$exitButton.Margin = 5
$exitButton.Add_Click({
    $window.Close()
})

$mainContainer.Children.Add($exitButton)

# Set the main container as the window content
$window.Content = $mainContainer

# Call the Refresh-Scripts function to load scripts on first launch
Refresh-Scripts

# Show the window
$window.ShowDialog()

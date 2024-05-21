Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$vmwarePath = Join-Path -Path $PSScriptRoot -ChildPath "Vmware"
$adFolder = Join-Path -Path $PSScriptRoot -ChildPath "AD"

# Ensure the VMware directory exists
if (-not (Test-Path $vmwarePath)) {
    [System.Windows.Forms.MessageBox]::Show("VMware directory not found. Run Update CCS Toolset.")
}
# Ensure the AD directory exists
if (-not (Test-Path $adFolder)) {
    [System.Windows.Forms.MessageBox]::Show("AD directory not found. Run Update CCS Toolset.")
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'CCS Tools Launcher'
$form.Size = New-Object System.Drawing.Size(390, 250)  # Adjusted form size for less space
$form.StartPosition = 'CenterScreen'

# Function to create a centered button
function Create-Button {
    param ($text, $location, $size, $action)
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($location[0], $location[1])
    $button.Size = New-Object System.Drawing.Size($size[0], $size[1])
    $button.Text = $text
    $button.Add_Click($action)
    $form.Controls.Add($button)
}

# Check if modules are installed and display status
$importExcelInstalled = Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue
$powerCLIInstalled = Get-Module -ListAvailable -Name VMware.PowerCLI -ErrorAction SilentlyContinue

# Label for ImportExcel installation status
$statusTextExcel = New-Object System.Windows.Forms.Label
$statusTextExcel.Location = New-Object System.Drawing.Point(10, 55)
$statusTextExcel.Size = New-Object System.Drawing.Size(180, 20)
$statusTextExcel.Text = "ImportExcel: $(if ($importExcelInstalled) {'Installed'} else {'Not Installed'})"
$form.Controls.Add($statusTextExcel)

# Label for PowerCLI installation status
$statusTextPowerCLI = New-Object System.Windows.Forms.Label
$statusTextPowerCLI.Location = New-Object System.Drawing.Point(210, 55)
$statusTextPowerCLI.Size = New-Object System.Drawing.Size(180, 20)
$statusTextPowerCLI.Text = "PowerCLI: $(if ($powerCLIInstalled) {'Installed'} else {'Not Installed'})"
$form.Controls.Add($statusTextPowerCLI)

# Button dimensions and positions
$buttonWidth = 150
$buttonHeight = 23
$horizontalPadding = 10
$verticalPosition = 10

# VMware Toolset button
Create-Button -text 'VMware Toolset' -location (10, $verticalPosition) -size ($buttonWidth, $buttonHeight, $horizontalPadding) -action {
    $scriptPath = Join-Path -Path $vmwarePath -ChildPath "vmware_launcher_new.ps1"
    & $scriptPath
}

# AD Toolset button
Create-Button -text 'AD Toolset' -location (210, $verticalPosition) -size ($buttonWidth, $buttonHeight, $horizontalPadding) -action {
    $scriptPath = Join-Path -Path $adFolder -ChildPath "AD_launcher.ps1"
    & $scriptPath
}

$verticalPosition = 80  # Adjusted position for the install buttons

# Install ImportExcel button
Create-Button -text 'Install ImportExcel' -location (10, $verticalPosition) -size ($buttonWidth, $buttonHeight, $horizontalPadding) -action {
    try {
        Install-Module -Name ImportExcel -Force
        [System.Windows.Forms.MessageBox]::Show("ImportExcel installed successfully", "Info")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to install ImportExcel", "Error")
    }
}

# Install PowerCLI button
Create-Button -text 'Install PowerCLI' -location (210, $verticalPosition) -size ($buttonWidth, $buttonHeight, $horizontalPadding) -action {
    try {
        Install-Module -Name VMware.PowerCLI -Confirm:$false -Force -AllowClobber
        [System.Windows.Forms.MessageBox]::Show("PowerCLI installed successfully", "Info")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to install PowerCLI", "Error")
    }
}

# Position for bottom buttons
$verticalPosition = $form.Height - 100


# Show the form
$form.ShowDialog()

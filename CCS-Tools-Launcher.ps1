Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$vmwarePath = Join-Path -Path $PSScriptRoot -ChildPath "Vmware"
$adFolder = Join-Path -Path $PSScriptRoot -ChildPath "AD"
$gitPath = Join-Path -Path $PSScriptRoot -ChildPath "Git\bin\git.exe"

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
$form.Size = New-Object System.Drawing.Size(390,250)  # Adjusted form size for less space
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

# Update Tools button
Create-Button -text 'Update CCS Toolset' -location (110, $verticalPosition) -size ($buttonWidth, $buttonHeight, $horizontalPadding) -action {
    try {
        $repositoryBaseUrl = "https://github.com/wiegz15/CCS-Tools"
        $destinationPath = $PSScriptRoot
        $launcherScript = Join-Path -Path $destinationPath -ChildPath "CCS-Tools-Launcher.ps1"
        $backupPath = "$launcherScript.backup"

        # Backup the launcher script if it exists
        if (Test-Path $launcherScript) {
            Copy-Item -Path $launcherScript -Destination $backupPath -Force
        }

        # Perform git operations using the portable Git
        if (Test-Path (Join-Path -Path $destinationPath -ChildPath ".git")) {
            Set-Location $destinationPath
            & $gitPath fetch --all

            # Ensure you are on the Testing branch
            & $gitPath checkout Testing

            # Explicitly update the AD and VMware folders
            & $gitPath checkout origin/Testing -- AD
            & $gitPath checkout origin/Testing -- Vmware

            # Clean untracked files in these folders only
            & $gitPath clean -fdx AD
            & $gitPath clean -fdx Vmware

        } else {
            & $gitPath clone -b Testing --single-branch $repositoryBaseUrl $destinationPath
            Set-Location $destinationPath
            # Ensure only AD and VMware folders are present
            Remove-Item -Recurse -Exclude AD, Vmware, ".git", $(Split-Path $launcherScript -Leaf)
        }

        # Restore the launcher script from backup
        if (Test-Path $backupPath) {
            Copy-Item -Path $backupPath -Destination $launcherScript -Force
            Remove-Item -Path $backupPath
        }

        [System.Windows.Forms.MessageBox]::Show("AD and VMware tools updated successfully from the Testing branch", "Info")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to update AD and VMware tools", "Error")
    }
}

# Show the form
$form.ShowDialog()

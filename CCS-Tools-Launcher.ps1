Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$vmwarePath = Join-Path -Path $PSScriptRoot -ChildPath "Vmware"
$adFolder = Join-Path -Path $PSScriptRoot -ChildPath "AD"
$gitExe = "C:\Program Files\Git\cmd\git.exe"

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

# Function to check if winget is available
function Check-Winget {
    try {
        $wingetPath = (Get-Command winget -ErrorAction SilentlyContinue).Source
        if (-not $wingetPath) {
            Write-Host "winget not found. Please install winget and try again."
            throw "winget not found"
        }
        return $wingetPath
    } catch {
        Write-Host "Error checking winget: $_"
        throw $_
    }
}

# Function to install Git using winget
function Install-Git-With-Winget {
    try {
        $wingetPath = Check-Winget
        Write-Host "winget found at $wingetPath. Installing Git..."
        Start-Process -FilePath $wingetPath -ArgumentList "install --id Git.Git -e --source winget" -Wait
    } catch {
        Write-Host "Error installing Git with winget: $_"
        throw $_
    }
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

        Write-Host "Starting update process..."

        # Install Git if not already installed
        if (-not (Test-Path $gitExe)) {
            Write-Host "Git not found. Installing using winget..."
            Install-Git-With-Winget
        } else {
            Write-Host "Git already exists."
        }

        # Backup the launcher script if it exists
        if (Test-Path $launcherScript) {
            Copy-Item -Path $launcherScript -Destination $backupPath -Force
        }

        # Perform git operations using the installed Git
        if (Test-Path (Join-Path -Path $destinationPath -ChildPath ".git")) {
            Set-Location $destinationPath
            & $gitExe fetch --all
            Write-Host "Fetched all branches"

            # Ensure you are on the Testing branch
            & $gitExe checkout Testing
            Write-Host "Checked out Testing branch"

            # Explicitly update the AD and VMware folders
            & $gitExe checkout origin/Testing -- AD
            & $gitExe checkout origin/Testing -- Vmware
            Write-Host "Checked out AD and VMware folders"

            # Clean untracked files in these folders only
            & $gitExe clean -fdx AD
            & $gitExe clean -fdx Vmware
            Write-Host "Cleaned untracked files in AD and VMware folders"

        } else {
            Write-Host "Cloning repository..."
            & $gitExe clone -b Testing --single-branch $repositoryBaseUrl $destinationPath
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
        Write-Host "Error during update process: $_"
        [System.Windows.Forms.MessageBox]::Show("Failed to update AD and VMware tools", "Error")
    }
}

# Show the form
$form.ShowDialog()

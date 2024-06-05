# Define the URL for the Git portable download
$gitPortableUrl = "https://github.com/git-for-windows/git/releases/download/v2.40.1.windows.1/PortableGit-2.40.1-64-bit.7z.exe"
$scriptPath = $PSScriptRoot
$gitPortablePath = "$scriptPath\GitPortable"

# Define repository URL and local repository path
$repositoryUrl = "https://github.com/wiegz15/CCS-Tools"
$localRepoPath = "$($scriptPath)\..\YourLocalRepo" -replace "\\", "/"

# Paths for AD and VMware folders
$adFolderPath = "$localRepoPath\AD"
$vmwareFolderPath = "$localRepoPath\VMware"

# Function to download and extract Git portable
function Install-GitPortable {
    if (-Not (Test-Path $gitPortablePath)) {
        Write-Output "Downloading Git portable..."
        $downloadPath = "$scriptPath\PortableGit.7z.exe"
        Invoke-WebRequest -Uri $gitPortableUrl -OutFile $downloadPath

        Write-Output "Extracting Git portable..."
        & $downloadPath -o"$scriptPath\GitPortable" -y
        Remove-Item $downloadPath -Force
    }
}

# Function to check for changes in AD and VMware folders
function Check-RepositoryChanges {
    Write-Output "Checking for changes in repository..."

    # Add GitPortable to PATH temporarily
    $env:PATH = "$gitPortablePath\cmd;$env:PATH"

    if (-Not (Test-Path $localRepoPath)) {
        Write-Output "Cloning repository..."
        git clone $repositoryUrl $localRepoPath
    } else {
        Set-Location $localRepoPath
        git pull
    }

    Write-Output "Checking for changes in AD folder..."
    $adStatus = git status $adFolderPath
    Write-Output $adStatus

    Write-Output "Checking for changes in VMware folder..."
    $vmwareStatus = git status $vmwareFolderPath
    Write-Output $vmwareStatus
}

# Main script execution
Install-GitPortable
Check-RepositoryChanges

Write-Output "Script execution completed."

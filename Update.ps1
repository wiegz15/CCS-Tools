# Define variables
$gitUrl = "https://github.com/git-for-windows/git/releases/download/v2.39.2.windows.1/PortableGit-2.39.2-64-bit.7z.exe"
$scriptPath = $PSScriptRoot
$gitFolderPath = Join-Path $scriptPath "PortableGit"
$gitExecutablePath = Join-Path $gitFolderPath "cmd\git.exe"
$gitArchivePath = Join-Path $scriptPath "PortableGit.7z.exe"
$repositoryUrl = "https://github.com/wiegz15/CCS-Tools/"
$localRepoPath = $scriptPath

# Create the folder for Git Portable if it doesn't exist
if (-Not (Test-Path -Path $gitFolderPath)) {
    New-Item -ItemType Directory -Path $gitFolderPath
}

# Download Git Portable if it is not already present
if (-Not (Test-Path -Path $gitExecutablePath)) {
    Write-Host "Downloading Git Portable..."
    Invoke-WebRequest -Uri $gitUrl -OutFile $gitArchivePath

    # Extract the archive to the specified folder
    Write-Host "Extracting Git Portable..."
    Start-Process -FilePath $gitArchivePath -ArgumentList "-o$gitFolderPath -y" -NoNewWindow -Wait

    # Clean up the downloaded archive
    Remove-Item $gitArchivePath
} else {
    Write-Host "Git Portable is already present."
}

# Update the PATH environment variable to include the Git Portable path
$gitDir = $gitExecutablePath.DirectoryName
$envPath = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::User)
if ($envPath -notlike "*$gitDir*") {
    $newPath = "$envPath;$gitDir"
    [System.Environment]::SetEnvironmentVariable("Path", $newPath, [System.EnvironmentVariableTarget]::User)
    Write-Host "Updated PATH environment variable to include Git Portable."
} else {
    Write-Host "Git Portable is already in the PATH environment variable."
}

# Use Git Portable to update the repository
if (-Not (Test-Path -Path (Join-Path $localRepoPath ".git"))) {
    # Clone the repository if it doesn't exist locally
    Write-Host "Cloning repository from $repositoryUrl to $localRepoPath..."
    & "$gitExecutablePath" clone $repositoryUrl $localRepoPath
} else {
    # Pull the latest changes if the repository already exists locally
    Write-Host "Updating repository at $localRepoPath..."
    Set-Location -Path $localRepoPath
    & "$gitExecutablePath" pull
}

Write-Host "Git repository is up to date."

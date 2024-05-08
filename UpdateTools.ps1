$repositoryBaseUrl = "https://github.com/wiegz15/CCS-Tools"
$destinationPath = Join-Path -Path $PSScriptRoot -ChildPath "App"

# Check if the destination directory already exists
if (Test-Path $destinationPath) {
    # If it exists, navigate to the directory and perform a git pull to update
    Set-Location $destinationPath
    git pull
} else {
    # If it doesn't exist, clone the repository
    git clone $repositoryBaseUrl $destinationPath
}

Write-Host "Update complete."
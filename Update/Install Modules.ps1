# Define all required modules
$modules = 'ImportExcel', 'VMware.PowerCLI'

# Function to check if a module is installed
function Is-ModuleInstalled {
    param (
        [string]$ModuleName
    )
    $module = Get-Module -ListAvailable -Name $ModuleName
    return $module -ne $null
}

# Function to check and install the latest NuGet provider
function Ensure-NuGetProvider {
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet)) {
        Write-Host "NuGet provider not found. Installing..."
        Install-PackageProvider -Name NuGet -Force -ErrorAction Stop
    } else {
        Write-Host "NuGet provider is already installed."
    }
}

try {
    Ensure-NuGetProvider
    Write-Host "NuGet provider installed or already present."
} catch {
    Write-Error "Failed to install NuGet provider. Error details: $_"
    exit
}

foreach ($module in $modules) {
    if (-not (Is-ModuleInstalled -ModuleName $module)) {
        try {
            # Attempt to install the module
            Install-Module -Name $module -Force -AllowClobber
            Write-Host "Module '$module' installed successfully."
        } catch {
            Write-Error "Failed to install module '$module'. Error details: $_"
        }
    } else {
        Write-Host "Module '$module' is already installed."
    }
}

Write-Host "Script execution completed."

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

foreach ($module in $modules) {
    if (-not (Is-ModuleInstalled -ModuleName $module)) {
        try {
            # Attempt to install the module
            Install-Module -Name $module -Force -Scope CurrentUser
            Write-Host "Module '$module' installed successfully."
        } catch {
            Write-Error "Failed to install module '$module'. Error details: $_"
        }
    } else {
        Write-Host "Module '$module' is already installed."
    }
}

Write-Host "Script execution completed."

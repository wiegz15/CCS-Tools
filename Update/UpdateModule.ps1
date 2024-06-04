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

# Iterate over each module and check if it is installed
foreach ($module in $modules) {
    if (-not (Is-ModuleInstalled -ModuleName $module)) {
        # If the module is not installed, prompt the user to install it
        $userInput = Read-Host "Module '$module' is not installed. Do you want to install it now? (y/n)"
        if ($userInput -eq 'y') {
            try {
                # Attempt to install the module
                Install-Module -Name $module -Force -Scope CurrentUser
                Write-Host "Module '$module' installed successfully."
            } catch
                        {
                Write-Error "Failed to install module '$module'. Error details: $_"
            }
        } else {
            Write-Host "Skipping installation of module '$module'."
        }
    } else {
        Write-Host "Module '$module' is already installed."
    }
}

Write-Host "Script execution completed."
# Ensure that the DHCP Server PowerShell module is loaded
Import-Module DhcpServer

# Function to check and install ImportExcel module if not present
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    
    # Check if the module is already installed
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName module is not installed. Installing now..."
        try {
            Install-Module -Name $ModuleName -Force -Scope CurrentUser -ErrorAction Stop
            Write-Host "$ModuleName module installed successfully."
        }
        catch {
            Write-Host "Failed to install $ModuleName. Error: $_"
            exit 1
        }
    } else {
        Write-Host "$ModuleName module is already installed."
    }
}

# Function to convert IP address to an integer
function ConvertTo-Int {
    param (
        [IPAddress]$ip
    )
    # Convert each octet of the IP address to an integer
    $bytes = $ip.GetAddressBytes()
    return [bitconverter]::ToUInt32($bytes[3..0], 0)
}

# Specify the output file location
$outputFilePath = "C:\scripts\DhcpScopeInfo.xlsx"

# Initialize an array to store DHCP scope information
$scopeInfo = @()

# Get all DHCP scopes
$scopes = Get-DhcpServerv4Scope

foreach ($scope in $scopes) {
    # Get scope details
    $scopeID = $scope.ScopeId
    $totalAddresses = (ConvertTo-Int $scope.EndRange) - (ConvertTo-Int $scope.StartRange) + 1
    $leasedAddresses = (Get-DhcpServerv4Lease -ScopeId $scopeID).Count
    $availableAddresses = $totalAddresses - $leasedAddresses
    
    # Create a custom object with the required information
    $scopeInfo += [pscustomobject]@{
        "ScopeName"        = $scope.Name
        "ScopeIPRange"     = "$($scope.StartRange) - $($scope.EndRange)"
        "TotalIPAddresses" = $totalAddresses
        "IssuedIPAddresses"= $leasedAddresses
        "AvailableIPs"     = $availableAddresses
    }
}

# Export the data to an Excel file
$scopeInfo | Export-Excel -Path $outputFilePath -AutoSize -Title "DHCP Scope Information"

# Output location of the Excel file
Write-Host "DHCP Scope information exported to: $outputFilePath"

# Ensure that the DHCP Server PowerShell module is loaded
Import-Module DhcpServer

# Function to convert IP address to an integer
function ConvertTo-Int {
    param (
        [IPAddress]$ip
    )
    $bytes = $ip.GetAddressBytes()
    return [bitconverter]::ToUInt32($bytes[3..0], 0)
}

# Function to get DHCP servers in the domain
function Get-DomainDHCPServers {
    $dhcpServers = Get-DhcpServerInDC
    return $dhcpServers | Select-Object DNSName, IPAddress
}

# Function to get DHCP scope information
function Get-DHCPScopeInfo {
    param (
        [string]$DHCPServer
    )
    
    # Initialize an array to store DHCP scope information
    $scopeInfo = @()

    # Get all DHCP scopes from the specified server
    $scopes = Get-DhcpServerv4Scope -ComputerName $DHCPServer

    foreach ($scope in $scopes) {
        # Get scope details
        $scopeID = $scope.ScopeId
        $totalAddresses = (ConvertTo-Int $scope.EndRange) - (ConvertTo-Int $scope.StartRange) + 1
        $leasedAddresses = (Get-DhcpServerv4Lease -ScopeId $scopeID -ComputerName $DHCPServer).Count
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

    return $scopeInfo
}

# Main script
# Get DHCP servers in the domain
$dhcpServers = Get-DomainDHCPServers

# Display DHCP servers in Out-GridView and let user select one
$selectedServer = $dhcpServers | Out-GridView -Title "Select a DHCP Server" -OutputMode Single

if ($selectedServer) {
    $dhcpServerName = $selectedServer.DNSName

    # Get DHCP scope information
    $scopeInfo = Get-DHCPScopeInfo -DHCPServer $dhcpServerName

    $excelPath = Join-Path -Path $reportsDir -ChildPath "DHCP_Scope_Report.xlsx"

    # Export the data to an Excel file
    $scopeInfo | Export-Excel -Path $excelPath -WorksheetName "DHCP_Scopes" -AutoSize -TableName "DHCP_Scopes" -TableStyle Medium11 -Append

    Write-Host "DHCP scope report for server $dhcpServerName has been generated and saved to: $excelPath"
} else {
    Write-Host "No DHCP server was selected. Exiting script."
}
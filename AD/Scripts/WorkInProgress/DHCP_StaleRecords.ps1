# Load the ImportExcel module
Import-Module ImportExcel

# Retrieve all DHCP servers in the domain
$machines = Get-DhcpServerInDC

# Initialize an array to hold stale lease data
$staleData = @()

# Iterate through each DHCP server
foreach ($machine in $machines) {
    $DHCPName = $machine.DnsName
    
    # Retrieve all scopes for the current DHCP server
    $AllScopes = Get-DhcpServerv4Scope -ComputerName $DHCPName
    
    # Iterate through each scope
    foreach ($scope in $AllScopes) {
        $ScopeName = $scope.Name
        $ScopeId = $scope.ScopeId
        
        # Retrieve lease information
        $leases = Get-DhcpServerv4Lease -ScopeId $ScopeId -ComputerName $DHCPName
        
        # Iterate through leases to check for stale records
        foreach ($lease in $leases) {
            if ($lease.ExpiryTime -lt (Get-Date)) {
                # Create a custom object with the data
                $obj = [PSCustomObject]@{
                    DHCPServer      = $DHCPName
                    ScopeName       = $ScopeName
                    ScopeID         = $ScopeId
                    IPAddress       = $lease.IPAddress
                    ClientID        = $lease.ClientId
                    HostName        = $lease.HostName
                    LeaseExpiry     = $lease.ExpiryTime
                    IsStale         = $true
                }
                
                # Add the object to the stale data array
                $staleData += $obj
            }
        }
    }
}

# Path for the Excel file
$excelPath = Join-Path -Path $reportsDir -ChildPath "DHCP_Stale.xlsx"

# Export stale data to Excel
$staleData | Export-Excel -Path $excelPath -WorksheetName "Stale_Leases" -AutoSize -TableName "Stale_Leases" -TableStyle Medium11 -Append

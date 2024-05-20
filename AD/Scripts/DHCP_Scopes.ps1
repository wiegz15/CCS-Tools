# Load the ImportExcel module
Import-Module ImportExcel

# Retrieve all DHCP servers in the domain
$machines = Get-DhcpServerInDC

# Initialize an array to hold the data
$data = @()

# Iterate through each DHCP server
foreach ($machine in $machines) {
    $DHCPName = $machine.DnsName
    
    # Retrieve all scopes for the current DHCP server
    $AllScopes = Get-DhcpServerv4Scope -ComputerName $DHCPName
    
    # Iterate through each scope
    foreach ($scope in $AllScopes) {
        $ScopeName = $scope.Name
        $ScopeId = $scope.ScopeId
        
        # Retrieve scope statistics
        $ScopeStats = Get-DhcpServerv4ScopeStatistics -ScopeId $ScopeId -ComputerName $DHCPName
        
        # Gather the statistics
        $ScopePercentInUse = $ScopeStats.PercentageInUse
        $Addfree = $ScopeStats.AddressesFree
        $AddUse = $ScopeStats.AddressesInUse
        
        # Create a custom object with the data
        $obj = [PSCustomObject]@{
            DHCPServer      = $DHCPName
            ScopeName       = $ScopeName
            ScopeID         = $ScopeId
            FreeAddresses   = $Addfree
            AddressesInUse  = $AddUse
            PercentInUse    = $ScopePercentInUse
        }
        
        # Add the object to the array
        $data += $obj
    }
}


$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$data | Export-Excel -Path $excelPath -WorksheetName "DHCP_Scopes" -AutoSize -TableName "DHCP_Scopes" -TableStyle Medium11 -Append
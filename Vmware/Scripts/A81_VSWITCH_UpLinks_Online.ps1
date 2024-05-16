
# Retrieve all ESXi hosts
$esxiHosts = Get-VMHost

# Initialize an array to hold the results
$results = @()

foreach ($esxiHost in $esxiHosts) {
    # Retrieve physical network adapters for the current host
    $pNICs = Get-VMHostNetworkAdapter -VMHost $esxiHost -Physical

    foreach ($pNIC in $pNICs) {
        # Create a custom object for each NIC with relevant details
        $result = New-Object PSObject -Property @{
            HostName = $esxiHost.Name
            PhysicalNIC = $pNIC.Name
            MACAddress = $pNIC.Mac
            LinkStatus = $pNIC.LinkStatus
            Duplex = $pNIC.Duplex
            BitRatePerSec = $pNIC.BitRatePerSec
        }

        # Add the custom object to the results array
        $results += $result
    }
}

# Display the results in an Out-GridView
# $results | Out-GridView -Title "Actual Speed of Physical NICs Across Hosts"

# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "NIC_CONNECTED"

# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "NIC_CONNECTED"


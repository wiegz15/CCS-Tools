# Get all ESXi hosts
$esxiHosts = Get-VMHost

# Create an array to hold our results
$results = @()

# Loop through each host and get all VMkernel adapters
foreach ($esxiHost in $esxiHosts) {
    $vmkAdapters = Get-VMHostNetworkAdapter -VMHost $esxiHost -VMKernel

    foreach ($adapter in $vmkAdapters) {
        $results += [PSCustomObject]@{
            HostName = $esxiHost.Name
            AdapterName = $adapter.Name
            IPAddress = $adapter.IP
            SubnetMask = $adapter.SubnetMask
            Management = if ($adapter.ManagementTrafficEnabled) { "Yes" } else { "No" }
            vMotion = if ($adapter.VMotionEnabled) { "Yes" } else { "No" }
            FaultTolerance = if ($adapter.FaultToleranceLoggingEnabled) { "Yes" } else { "No" }
            vSAN = if ($adapter.VsanTrafficEnabled) { "Yes" } else { "No" }
        }
    }
}



# Display results in Out-GridView
# $results | Out-GridView -Title "ESXi Hosts VMkernel Adapter Information"

# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "HostIPAddresses"

# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName $worksheetName


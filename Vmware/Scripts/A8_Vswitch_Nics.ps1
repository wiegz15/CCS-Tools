
# Retrieve all ESXi hosts
$esxiHosts = Get-VMHost

# Initialize an array to hold the results
$results = @()

foreach ($esxiHost in $esxiHosts) {
    # Retrieve all virtual switches for the current host
    $vSwitches = Get-VirtualSwitch -VMHost $esxiHost
    
    foreach ($vSwitch in $vSwitches) {
        # Retrieve physical network adapters connected to the current virtual switch
        $pNICs = $vSwitch.ExtensionData.Spec.Bridge.NicDevice
        if ($pNICs -eq $null) {
            $pNICs = 'None'
        }

        $result = New-Object PSObject -Property @{
            HostName = $esxiHost.Name
            VirtualSwitch = $vSwitch.Name
            PhysicalNICs = $pNICs -join ', '
        }

        # Add the custom object to the results array
        $results += $result
    }
}

# Display the results in an Out-GridView
# $results | Out-GridView -Title "Virtual Switches and Attached Hosts"


# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "VSWITCHESAttachedHosts"

# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "VSWITCHESAttachedHosts"


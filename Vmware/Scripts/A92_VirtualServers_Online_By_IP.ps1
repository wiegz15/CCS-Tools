
# Get all VMs
$vms = Get-VM

# Prepare data to display
$vmInfo = foreach ($vm in $vms) {
    # Check if VMware Tools is installed and running to get the IP Address, and also check the power state
    $ipAddress = $vm.ExtensionData.Guest.IpAddress
    $powerState = $vm.PowerState
    $hostServer = $vm.VMHost.Name
    $ipAddressOrStatus = if ($powerState -eq "PoweredOn" -and $ipAddress) {
        $ipAddress -join ', ' # Join multiple IP addresses with comma if more than one
    } elseif ($powerState -eq "PoweredOff") {
        "VM Powered Off"
    } else {
        "IP Not Available or VMware Tools not running"
    }
    
    [PSCustomObject]@{
        VMName = $vm.Name
        HostServer = $hostServer
        PowerState = $powerState
        IPAddressOrStatus = $ipAddressOrStatus
    }
}

# Display in Out-GridView
# $vmInfo | Out-GridView -Title "VM Details"


# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "VM_Online"

# Export the results to an Excel file
$vminfo | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "VM_Online"


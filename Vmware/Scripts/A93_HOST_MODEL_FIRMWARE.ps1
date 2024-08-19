# Initialize an array to store host information
$hostInfoArray = @()

# Get all ESXi hosts
$allHosts = Get-VMHost

# Loop through each host and collect make, model, serial number, firmware levels, total memory, total CPU, and cluster name
foreach ($vmHost in $allHosts) {
    $hostView = Get-View -ViewType HostSystem -Property Name, Parent, Hardware.SystemInfo, Hardware.BiosInfo, Hardware.CpuInfo, Hardware.MemorySize -Filter @{"Name"=$vmHost.Name}
    $make = $hostView.Hardware.SystemInfo.Vendor
    $model = $hostView.Hardware.SystemInfo.Model
    $serialNumber = $hostView.Hardware.SystemInfo.OtherIdentifyingInfo | Where-Object {$_.IdentifierType.Key -eq "ServiceTag"} | Select-Object -ExpandProperty IdentifierValue
    $firmware = $hostView.Hardware.BiosInfo.BiosVersion
    $totalMemoryGB = [math]::round($hostView.Hardware.MemorySize / 1GB, 2)
    $totalCpuCores = $hostView.Hardware.CpuInfo.NumCpuCores
    
    # Get the cluster name
    $parent = Get-View $hostView.Parent
    $clusterName = $parent.Name

    # Create a custom object for each host with the desired information
    $hostInfo = New-Object PSObject -Property @{
        Make          = $make
        Model         = $model
        SerialNumber  = $serialNumber
        Firmware      = $firmware
        Name          = $vmHost.Name
        TotalMemoryGB = $totalMemoryGB
        TotalCpuCores = $totalCpuCores
        Cluster       = $clusterName
    }

    # Add the host info to the array
    $hostInfoArray += $hostInfo
}

# Explicitly select properties in the desired order and export to Excel
$hostInfoArray | Select-Object Make, Model, SerialNumber, Firmware, Name, TotalMemoryGB, TotalCpuCores, Cluster | Export-Excel -Path $excelpath -WorksheetName "HostModelFirmwareCapacity"


# Export the results to an Excel file
#$hostInfoArray | Export-Excel -Path $excelPath -WorksheetName "HostInventory" -AutoSize -TableName "HostInventory"




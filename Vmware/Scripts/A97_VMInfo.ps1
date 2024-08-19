# Get all VMs
$vms = Get-VM

# Create an array to store the results
$results = @()

foreach ($vm in $vms) {
    $networkInfo = $vm | Get-NetworkAdapter | Select-Object -ExpandProperty NetworkName
    $ipAddress = $vm.Guest.IPAddress | Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+' } | Select-Object -First 1
    $datastores = ($vm | Get-Datastore | Select-Object -ExpandProperty Name) -join ', '
    
    $result = [PSCustomObject]@{
        Name = $vm.Name
        MemoryGB = $vm.MemoryGB
        NumCPU = $vm.NumCpu
        Networks = ($networkInfo -join ', ')
        StorageGB = [math]::Round(($vm | Get-HardDisk | Measure-Object -Property CapacityGB -Sum).Sum, 2)
        Datastores = $datastores
        IPAddress = if ($ipAddress) { $ipAddress } else { "N/A" }
        VMToolsVersion = $vm.Guest.ToolsVersion
        GuestOS = $vm.Guest.OSFullName
        PowerState = $vm.PowerState
        HardwareVersion = $vm.HardwareVersion
    }
    
    $results += $result
}


# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName "VMInfo" -AutoSize -TableName "VMInfo"


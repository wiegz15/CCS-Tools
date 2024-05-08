
# Collect info for each host
$hostsInfo = @()

# Retrieve all datastores and map them to their corresponding storage device (LUN)
$datastoreDeviceMap = @{}
Get-Datastore | ForEach-Object {
    $ds = $_
    $ds.ExtensionData.Info.Vmfs.Extent | ForEach-Object {
        $datastoreDeviceMap[$_.DiskName] = $ds.Name
    }
}

Get-VMHost | ForEach-Object {
    $esxHost = $_
    $esxcli = Get-EsxCli -VMHost $esxHost
    $storagePaths = $esxcli.storage.core.path.list()
    
    # Group paths by device and filter for FC (Fibre Channel) transport only
    $activeFCPathsPerDevice = $storagePaths | Where-Object { $_.Transport -eq "FC" } | Group-Object -Property Device | ForEach-Object {
        $activePathsCount = ($_ | Select-Object -ExpandProperty Group | Where-Object { $_.State -eq "active" }).Count
        [PSCustomObject]@{
            Device = $_.Name
            ActiveFCPathsCount = $activePathsCount
        }
    }

    foreach ($item in $activeFCPathsPerDevice) {
        # Attempt to find the datastore name using the device mapping
        $datastoreName = $datastoreDeviceMap[$item.Device]
        if ($datastoreName) {
            $hostsInfo += [PSCustomObject]@{
                HostName = $esxHost.Name
                StorageDeviceName = $item.Device
                DatastoreName = $datastoreName
                ActiveFCConnectionsCount = $item.ActiveFCPathsCount
            }
        }
    }
}

# Display the collected information
# $hostsInfo | Out-GridView



# Import the ImportExcel module
# Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "DS_MPath"

# Export the results to an Excel file
$hostsInfo | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table3"


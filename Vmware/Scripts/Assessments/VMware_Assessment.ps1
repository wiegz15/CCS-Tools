Import-Module ImportExcel

# Determine the path to the vSphereOvercommit module
$scriptDir = Split-Path -Parent $PSCommandPath
$modulePath = Join-Path -Path (Split-Path -Parent (Split-Path -Parent $scriptDir)) -ChildPath "bin\vSphereOvercommit.psm1"

# Import the vSphereOvercommit module
Import-Module $modulePath

# Get all the clusters
$clusters = Get-Cluster
$clusterData = @()
$hostData = @()
$averageData = @()
$vmData = @()
$datastoreData = @()

# Loop through the clusters and extract the required data
foreach ($cluster in $clusters) {
    Write-Host "Processing cluster: $($cluster.Name)"
    $vmhosts = Get-VMHost -Location $cluster
    $num_esxis = $vmhosts.Count
     
    # Calculate cluster metrics
    $ClusterPoweredOnvCPUs = (Get-VM -Location $cluster | Where-Object {$_.PowerState -eq "PoweredOn" }).NumCpu | Measure-Object -Sum
    $ClusterCPUCores = ($vmhosts | Measure-Object -Property NumCpu -Sum).Sum

    $ClusterPoweredOnvRAM = (Get-VM -Location $cluster | Where-Object {$_.PowerState -eq "PoweredOn" }).MemoryGB | Measure-Object -Sum
    $ClusterPhysRAM = ($vmhosts | Measure-Object -Property MemoryTotalGB -Sum).Sum

    $oneesxi = if ($num_esxis) { $num_esxis - 1 } else { $null }
    $coresPerESXi = if ($num_esxis) { $ClusterCPUCores / $num_esxis } else { $null }
    $coresPerClusterMinusOne = $coresPerESXi * $oneesxi
    $TotalCoresPerClusterMinusOne = [Math]::Round($coresPerClusterMinusOne)

    $ramPerESXi = $ClusterPhysRAM / $num_esxis
    $ramPerClusterMinusOne = $ramPerESXi * $oneesxi
    $TotalRAMPerClusterMinusOne = [Math]::Round($ramPerClusterMinusOne)

    # Retrieve average CPU and RAM usage over the past day
    $avgCPUUsage = Get-Stat -Entity $vmhosts -Stat "cpu.usage.average" -Start (Get-Date).AddDays(-7) -IntervalMins 30 | Measure-Object -Property Value -Average
    $avgRAMUsage = Get-Stat -Entity $vmhosts -Stat "mem.usage.average" -Start (Get-Date).AddDays(-7) -IntervalMins 30 | Measure-Object -Property Value -Average

    $avgCPUSum = $avgCPUUsage.Average * $ClusterCPUCores / 100
    $avgRAMSum = $avgRAMUsage.Average * $ClusterPhysRAM / 100

    $avgCoresPerClusterMinusOne = $coresPerESXi * $oneesxi
    $TotalAvgCoresPerClusterMinusOne = [Math]::Round($avgCoresPerClusterMinusOne)

    $avgRAMPerClusterMinusOne = $ramPerESXi * $oneesxi
    $TotalAvgRAMPerClusterMinusOne = [Math]::Round($avgRAMPerClusterMinusOne)

    # Create an ordered hashtable for each cluster
    $clusterProperty = [ordered]@{
        "vCenter" = $vcenter
        "Cluster Name" = $cluster.Name
        "Number of ESXi servers" = $num_esxis
        "pCPU" = $ClusterCPUCores
        "vCPU" = (Get-VM -Location $cluster | Measure-Object -Property NumCpu -Sum).Sum
        "PoweredOn vCPUs" = $ClusterPoweredOnvCPUs.Sum
        "PoweredOn VMs" = (Get-VM -Location $cluster | Where-Object { $_.PowerState -eq "PoweredOn" }).Count
        "vCPU:pCPU Ratio" = if ($ClusterCPUCores -ne 0) { [Math]::Round($ClusterPoweredOnvCPUs.Sum / $ClusterCPUCores, 3) } else { 0 }
        "vCPU:pCPU Ratio with one ESXi failed" = if ($TotalCoresPerClusterMinusOne -ne 0) { [Math]::Round($ClusterPoweredOnvCPUs.Sum / $TotalCoresPerClusterMinusOne, 3) } else { 0 }
        "CPU Overcommit (%)" = if ($ClusterCPUCores -ne 0) { [Math]::Round(100 * (($ClusterPoweredOnvCPUs.Sum - $ClusterCPUCores) / $ClusterCPUCores), 2) } else { 0 }
        "pRAM(GB)" = $ClusterPhysRAM
        "vRAM(GB)" = [Math]::Round((Get-VM -Location $cluster | Measure-Object -Property MemoryGB -Sum).Sum, 2)
        "PoweredOn vRAM (GB)" = $ClusterPoweredOnvRAM.Sum
        "vRAM:pRAM Ratio" = if ($ClusterPhysRAM -ne 0) { [Math]::Round($ClusterPoweredOnvRAM.Sum / $ClusterPhysRAM, 3) } else { 0 }
        "vRAM:pRAM Ratio with one ESXi failed" = if ($TotalRAMPerClusterMinusOne -ne 0) { [Math]::Round($ClusterPoweredOnvRAM.Sum / $TotalRAMPerClusterMinusOne, 3) } else { 0 }
        "RAM Overcommit (%)" = if ($ClusterPhysRAM -ne 0) { [Math]::Round(100 * (($ClusterPoweredOnvRAM.Sum - $ClusterPhysRAM) / $ClusterPhysRAM), 2) } else { 0 }
        "Avg CPU Usage (%)" = [Math]::Round($avgCPUUsage.Average, 2)
        "Avg RAM Usage (%)" = [Math]::Round($avgRAMUsage.Average, 2)
    }

    # Add the cluster data to the array
    $clusterData += New-Object -TypeName psobject -Property $clusterProperty

    # Create an ordered hashtable for average data
    $averageProperty = [ordered]@{
        "vCenter" = $vcenter
        "Cluster Name" = $cluster.Name
        "Avg CPU Usage (%)" = [Math]::Round($avgCPUUsage.Average, 2)
        "Avg RAM Usage (%)" = [Math]::Round($avgRAMUsage.Average, 2)
        "Avg vCPU:pCPU Ratio" = if ($ClusterCPUCores -ne 0) { [Math]::Round($avgCPUSum / $ClusterCPUCores, 3) } else { 0 }
        "Avg vCPU:pCPU Ratio with one ESXi failed" = if ($TotalAvgCoresPerClusterMinusOne -ne 0) { [Math]::Round($avgCPUSum / $TotalAvgCoresPerClusterMinusOne, 3) } else { 0 }
        "Avg RAM Overcommit (%)" = if ($ClusterPhysRAM -ne 0) { [Math]::Round(100 * (($avgRAMSum - $ClusterPhysRAM) / $ClusterPhysRAM), 2) } else { 0 }
        "Avg vRAM:pRAM Ratio" = if ($ClusterPhysRAM -ne 0) { [Math]::Round($avgRAMSum / $ClusterPhysRAM, 3) } else { 0 }
        "Avg vRAM:pRAM Ratio with one ESXi failed" = if ($TotalAvgRAMPerClusterMinusOne -ne 0) { [Math]::Round($avgRAMSum / $TotalAvgRAMPerClusterMinusOne, 3) } else { 0 }
    }

    # Add the average data to the array
    $averageData += New-Object -TypeName psobject -Property $averageProperty

    # Loop through each host in the cluster to gather host data
    foreach ($vmHost in $vmhosts) {
        Write-Host "Processing host: $($vmHost.Name)"
        $hostView = Get-View $vmHost
        $hostModel = $hostView.Hardware.SystemInfo.Model
        $cpuCores = $hostView.Hardware.CpuInfo.NumCpuCores
        $ram = [Math]::Round($hostView.Hardware.MemorySize / 1GB, 2)
        $TotalvmCount = (Get-VM -Location $vmHost).Count

        # Get firmware version
        $firmwareVersion = $hostView.Hardware.BiosInfo.BiosVersion

        # Add the host data to the array
        $hostData += [PSCustomObject]@{
            HostName = $vmHost.Name
            Model = $hostModel
            CPUCores = $cpuCores
            RAMGB = $ram
            VMCount = $TotalvmCount
            FirmwareVersion = $firmwareVersion
        }
    }

    # Gather powered-on VMs data
    $poweredOnVMs = Get-VM -Location $cluster | Where-Object { $_.PowerState -eq "PoweredOn" }
    foreach ($vm in $poweredOnVMs) {
        # Calculate average CPU and RAM usage over the last 7 days
        $avgVMCPUUsage = Get-Stat -Entity $vm -Stat "cpu.usage.average" -Start (Get-Date).AddDays(-7) -IntervalMins 30 | Measure-Object -Property Value -Average
        $avgVMRAMUsage = Get-Stat -Entity $vm -Stat "mem.usage.average" -Start (Get-Date).AddDays(-7) -IntervalMins 30 | Measure-Object -Property Value -Average

        # Add the VM data to the array
        $vmData += [PSCustomObject]@{
            VMName = $vm.Name
            vCPU = $vm.NumCpu
            RAMGB = $vm.MemoryGB
            StorageGB = [Math]::Round(($vm.HardDisks | Measure-Object -Property CapacityGB -Sum).Sum, 2)
            AvgCPUUsage = if ($avgVMCPUUsage.Average) { [Math]::Round($avgVMCPUUsage.Average, 2) } else { 0 }
            AvgRAMUsage = if ($avgVMRAMUsage.Average) { [Math]::Round($avgVMRAMUsage.Average, 2) } else { 0 }
        }
    }

    # Gather datastore data
    $datastores = Get-Cluster $cluster | Get-Datastore | Where-Object { $_.Name -notlike "*ISO*" -and $_.Name -notlike "*loca*" } | Sort-Object -Property FreeSpaceGB -Descending
    foreach ($datastore in $datastores) {
        # Add the datastore data to the array
        $datastoreData += [PSCustomObject]@{
            DatastoreName = $datastore.Name
            FreeSpaceGB = [Math]::Round($datastore.FreeSpaceGB, 2)
            CapacityGB = [Math]::Round($datastore.CapacityGB, 2)
        }
    }
}

# Export the cluster data to the first sheet in the Excel file
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Assessment.xlsx"
$clusterData | Export-Excel -Path $excelPath -WorksheetName "ClusterSummary" -AutoSize -TableName "ClusterSummary"

# Export the host data to the second sheet in the Excel file
$hostData | Export-Excel -Path $excelPath -WorksheetName "HostSummary" -AutoSize -TableName "HostSummary" -Append

# Export the average data to the third sheet in the Excel file
$averageData | Export-Excel -Path $excelPath -WorksheetName "AverageUsageSummary" -AutoSize -TableName "AverageUsageSummary" -Append

# Export the VM data to the fourth sheet in the Excel file
$vmData | Export-Excel -Path $excelPath -WorksheetName "PoweredOnVMs" -AutoSize -TableName "PoweredOnVMs" -Append

# Export the datastore data to the fifth sheet in the Excel file
$datastoreData | Export-Excel -Path $excelPath -WorksheetName "DatastoreSummary" -AutoSize -TableName "DatastoreSummary" -Append

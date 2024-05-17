Import-Module ImportExcel

# Get all the clusters
$clusters = Get-Cluster
$clusterData = @()
$hostData = @()

# Loop through the clusters and extract the required data
foreach ($cluster in $clusters) {
    Write-Host "Processing cluster: $($cluster.Name)"
    $vmhosts = Get-VMHost -Location $cluster
    $num_esxis = $vmhosts.Count
     
    # Calculate cluster metrics
    $ClusterPoweredOnvCPUs = (Get-VM -Location $cluster | Where-Object {$_.PowerState -eq "PoweredOn" }).NumCpu | Measure-Object -Sum
    $ClusterCPUCores = ($vmhosts.NumCpu | Measure-Object -Sum).Sum

    $ClusterPoweredOnvRAM = (Get-VM -Location $cluster | Where-Object {$_.PowerState -eq "PoweredOn" }).MemoryGB | Measure-Object -Sum
    $ClusterPhysRAM = ($vmhosts.MemoryTotalGB | Measure-Object -Sum).Sum

    $oneesxi = if ($num_esxis) { $num_esxis - 1 } else { $null }
    $coresPerESXi = if ($num_esxis) { $ClusterCPUCores / $num_esxis } else { $null }
    $coresPerClusterMinusOne = $coresPerESXi * $oneesxi
    $TotalCoresPerClusterMinusOne = [Math]::Round($coresPerClusterMinusOne)

    $ramPerESXi = $ClusterPhysRAM / $num_esxis
    $ramPerClusterMinusOne = $ramPerESXi * $oneesxi
    $TotalRAMPerClusterMinusOne = [Math]::Round($ramPerClusterMinusOne)

    # Create an ordered hashtable for each cluster
    $clusterProperty = [ordered]@{
        "vCenter" = $vcenter
        "Cluster Name" = $cluster.Name
        "Number of ESXi servers" = $num_esxis
        "pCPU" = $ClusterCPUCores
        "vCPU" = (Get-VM -Location $cluster | Measure-Object NumCpu -Sum).Sum
        "PoweredOn vCPUs" = $ClusterPoweredOnvCPUs.Sum
        "vCPU:pCPU Ratio" = [Math]::Round($ClusterPoweredOnvCPUs.Sum / $ClusterCPUCores, 3)
        "vCPU:pCPU Ratio with one ESXi failed" = [Math]::Round($ClusterPoweredOnvCPUs.Sum / $TotalCoresPerClusterMinusOne, 3)
        "CPU Overcommit (%)" = [Math]::Round(100 * ($ClusterPoweredOnvCPUs.Sum / $ClusterCPUCores), 3)
        'pRAM(GB)' = $ClusterPhysRAM
        'vRAM(GB)' = [Math]::Round((Get-VM -Location $cluster | Measure-Object MemoryGB -Sum).Sum, 2)
        'PoweredOn vRAM (GB)' = $ClusterPoweredOnvRAM.Sum
        'vRAM:pRAM Ratio' = [Math]::Round($ClusterPoweredOnvRAM.Sum / $ClusterPhysRAM, 3)
        "vRAM:pRAM Ratio with one ESXi failed" = [Math]::Round($ClusterPoweredOnvRAM.Sum / $TotalRAMPerClusterMinusOne, 3)
        'RAM Overcommit (%)' = [Math]::Round(100 * ($ClusterPoweredOnvRAM.Sum / $ClusterPhysRAM), 2)
    }

    # Add the cluster data to the array
    $clusterData += New-Object -TypeName psobject -Property $clusterProperty

    # Loop through each host in the cluster to gather host data
    foreach ($vmHost in $vmhosts) {
        Write-Host "Processing host: $($vmHost.Name)"
        $hostView = Get-View $vmHost
        $hostModel = $hostView.Hardware.SystemInfo.Model
        $cpuCores = $hostView.Hardware.CpuInfo.NumCpuCores
        $ram = [Math]::Round($hostView.Hardware.MemorySize / 1GB, 2)
        $vmCount = (Get-VM -Location $vmHost).Count
        
        # Get firmware version
        $firmwareVersion = $hostView.Hardware.BiosInfo.BiosVersion

        # Add the host data to the array
        $hostData += [PSCustomObject]@{
            HostName = $vmHost.Name
            Model = $hostModel
            CPUCores = $cpuCores
            RAMGB = $ram
            VMCount = $vmCount
            FirmwareVersion = $firmwareVersion
        }
    }
}

# Export the cluster data to the first sheet in the Excel file
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Assessment.xlsx"
$clusterData | Export-Excel -Path $excelPath -WorksheetName "Cluster Summary" -AutoSize -TableName "ClusterSummary"

# Export the host data to the second sheet in the Excel file
$hostData | Export-Excel -Path $excelPath -WorksheetName "Host Summary" -AutoSize -TableName "HostSummary" -Append


Import-Module ImportExcel

# Get all the clusters
$clusters = Get-Cluster
$data = @()

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
    $property = [ordered]@{
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
 
    # Add the hashtable to the data array
    $data += New-Object -TypeName psobject -Property $property
}
 
# Export the data to an Excel file
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Assessment.xlsx"
$data | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -TableName "Summary"

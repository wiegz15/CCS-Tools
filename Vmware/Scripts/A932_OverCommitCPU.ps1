# Get all clusters
$clusters = Get-Cluster

# Create an array to store the results
$results = @()

foreach ($cluster in $clusters) {
    $hosts = $cluster | Get-VMHost
    $totalPhysicalCPU = ($hosts | Measure-Object -Property NumCpu -Sum).Sum
    $totalCPUMhz = ($hosts | Measure-Object -Property CpuTotalMhz -Sum).Sum
    $totalUsedCPUMhz = ($hosts | Measure-Object -Property CpuUsageMhz -Sum).Sum
    $totalAllocatedCPU = ($cluster | Get-VM | Measure-Object -Property NumCpu -Sum).Sum

    $cpuOvercommitRatio = $totalAllocatedCPU / $totalPhysicalCPU
    $isOvercommitted = $cpuOvercommitRatio -gt 1

    $results += [PSCustomObject]@{
        ClusterName = $cluster.Name
        TotalPhysicalCPU = $totalPhysicalCPU
        TotalCPUMhz = $totalCPUMhz
        TotalUsedCPUMhz = $totalUsedCPUMhz
        TotalAllocatedCPU = $totalAllocatedCPU
        CPUOvercommitRatio = [math]::Round($cpuOvercommitRatio, 2)
        IsOvercommitted = $isOvercommitted
        CPUUsagePercentage = [math]::Round(($totalUsedCPUMhz / $totalCPUMhz) * 100, 2)
    }
}


# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName "OvercommitCPU" -AutoSize -TableName "OvercommitCPU"




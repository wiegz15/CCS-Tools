# Get all clusters
$clusters = Get-Cluster

# Create an array to store the results
$results = @()

foreach ($cluster in $clusters) {
    $hosts = $cluster | Get-VMHost
    $totalPhysicalMemory = ($hosts | Measure-Object -Property MemoryTotalGB -Sum).Sum
    $totalUsedMemory = ($hosts | Measure-Object -Property MemoryUsageGB -Sum).Sum
    $totalReservedMemory = ($hosts | Measure-Object -Property MemoryReservationGB -Sum).Sum
    $totalAllocatedMemory = ($cluster | Get-VM | Measure-Object -Property MemoryGB -Sum).Sum

    $overcommitRatio = $totalAllocatedMemory / ($totalPhysicalMemory - $totalReservedMemory)
    $isOvercommitted = $overcommitRatio -gt 1

    $results += [PSCustomObject]@{
        ClusterName = $cluster.Name
        TotalPhysicalMemoryGB = [math]::Round($totalPhysicalMemory, 2)
        TotalUsedMemoryGB = [math]::Round($totalUsedMemory, 2)
        TotalReservedMemoryGB = [math]::Round($totalReservedMemory, 2)
        TotalAllocatedMemoryGB = [math]::Round($totalAllocatedMemory, 2)
        OvercommitRatio = [math]::Round($overcommitRatio, 2)
        IsOvercommitted = $isOvercommitted
    }
}


# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName "OvercommitRAM" -AutoSize -TableName "OvercommitRam"




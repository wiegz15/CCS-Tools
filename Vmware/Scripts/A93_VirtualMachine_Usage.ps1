
# Define thresholds
$cpuThreshold = 80 # CPU utilization percentage
$memoryThreshold = 80 # Memory utilization percentage

# Retrieve all VM stats, excluding powered down VMs
$vms = Get-VM | Where-Object { $_.PowerState -eq "PoweredOn" } | ForEach-Object {
    $vmName = $_.Name
    $cpuUsage = ($_ | Get-Stat -Stat cpu.usage.average -Realtime -MaxSamples 1).Value
    $memoryUsage = ($_ | Get-Stat -Stat mem.usage.average -Realtime -MaxSamples 1).Value
    
    # Create a custom PSObject for each VM with relevant details
    [PSCustomObject]@{
        Name = $vmName
        CPUUsage = $cpuUsage
        MemoryUsage = $memoryUsage
    }
}

# Filter VMs exceeding thresholds
$exceedingVMs = $vms | Where-Object { $_.CPUUsage -gt $cpuThreshold -or $_.MemoryUsage -gt $memoryThreshold }

# Check if any VMs exceeded thresholds and display them
if ($exceedingVMs) {
   # $exceedingVMs | Select-Object Name, CPUUsage, MemoryUsage | Out-GridView -Title "VMs Exceeding Utilization Thresholds"

    # Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "HighUtilization_VMs"

# Export the results to an Excel file
$exceedingVMs | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table10"






} else {
    Write-Host "No VMs are exceeding utilization thresholds."
}


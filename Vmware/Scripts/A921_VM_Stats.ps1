# Retrieve VM information
$vms = Get-VM

# Create a custom object to store the results
$vmDetails = @()
$totalCPU = 0
$totalMemoryGB = 0

foreach ($vm in $vms) {
    $vmHost = $vm.VMHost.Name
    $memoryGB = [math]::Round($vm.MemoryMB / 1024, 2)
    $vmDetail = New-Object PSObject -Property @{
        Name      = $vm.Name
        CPU       = $vm.NumCpu
        MemoryGB  = $memoryGB
        HostServer = $vmHost
    }
    $vmDetails += $vmDetail
    $totalCPU += $vm.NumCpu
    $totalMemoryGB += $memoryGB
}

# Reorder properties for Out-GridView
$vmDetailsOrdered = $vmDetails | Select-Object Name, CPU, MemoryGB, HostServer


# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "VMStats"

# Export the results to an Excel file
$vmDetailsOrdered | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "VMStats"
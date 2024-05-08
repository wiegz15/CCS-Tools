# Define the critical services to check
$criticalServices = @("hostd", "vpxa", "ntpd", "TSM", "TSM-SSH")

# Get the list of all services and their current status, filtering for critical services, and display in Out-GridView
#Get-VMHost | ForEach-Object {
 #   $vmHost = $_.Name
  #  Get-VMHostService -VMHost $_ | Where-Object { $_.Key -in $criticalServices } | 
   # Select-Object @{Name="Host"; Expression={$vmHost}}, Key, Label, Running 
#} | Out-GridView -Title "Critical VMware ESXi Services Status"


$results = Get-VMHost | ForEach-Object {
    $vmHost = $_.Name
    Get-VMHostService -VMHost $_ | Where-Object { $_.Key -in $criticalServices } | 
    Select-Object @{Name="Host"; Expression={$vmHost}}, Key, Label, Running 
}

# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "hostservices"

# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table6"


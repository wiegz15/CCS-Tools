


# Retrieve all ESXi hosts and check for maintenance mode status
$esxiHosts = Get-VMHost | Select-Object Name, State, @{Name="InMaintenanceMode";Expression={$_.ConnectionState -eq "Maintenance"}}

# Display the results in an Out-GridView
#$esxiHosts | Out-GridView -Title "ESXi Hosts Maintenance Mode Status"



# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "Maint-Mode"

# Export the results to an Excel file
$esxiHosts | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Maint-Mode"



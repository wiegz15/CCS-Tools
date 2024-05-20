# Fetch admin permissions
$results = Get-VIPermission | Where-Object { $_.Role -eq "Admin" } | Select-Object Entity, Principal, Role

# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "AdminPermissions"

# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "AdminPermissions" -Append


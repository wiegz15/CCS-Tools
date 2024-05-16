# Retrieve all VMs and their hardware compatibility levels
$allVMs = Get-VM | Select-Object Name, HardwareVersion


# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "Vm_Hardware_Compat"

# Export the results to an Excel file
$allVMs | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Vm_Hardware_Compat"




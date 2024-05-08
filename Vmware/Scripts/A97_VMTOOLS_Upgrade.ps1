# Retrieve all VMs and check VMware Tools status
$vmList = Get-VM | Select-Object Name, @{N="VMwareToolsStatus";E={if($_.ExtensionData.Guest.ToolsStatus -eq "toolsOld"){"Old"}else{"Current"}}}

# Display the list in an Out-GridView
# $vmList | Out-GridView -Title "Virtual Machines and VMware Tools Status"




# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "VMTools_Updated"

# Export the results to an Excel file
$vmList | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table23"



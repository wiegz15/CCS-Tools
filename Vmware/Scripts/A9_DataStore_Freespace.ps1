# Import necessary modules - assume VMware.PowerCLI and ImportExcel modules are installed


# Search for datastores with less than 40 percent free space and display in a grid view
$datastoresWithLessThan40PercentFree = Get-Datastore | Where-Object {
    ($_.FreeSpaceGB / $_.CapacityGB) -lt 0.4
} | Select-Object Name, CapacityGB, FreeSpaceGB, @{Name="FreeSpacePercent"; Expression={[math]::Round(($_.FreeSpaceGB / $_.CapacityGB) * 100, 2)}}

# Display in Out-GridView
# $datastoresWithLessThan40PercentFree | Out-GridView -Title "Datastores with less than 40% free space"

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "Datastore_Freespace"

# Export the results to an Excel file
$datastoresWithLessThan40PercentFree | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table1"



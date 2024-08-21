
# Import the Group Policy module
Import-Module GroupPolicy

# Verify if the Group Policy module was imported successfully
if (-not (Get-Module -Name GroupPolicy)) {
    Write-Error "Failed to import GroupPolicy module. Please ensure it is installed correctly."
    exit
}

# Get all GPOs in the domain
$gpos = Get-GPO -All

# Check if any GPOs were retrieved
if ($null -eq $gpos) {
    Write-Error "Failed to retrieve GPOs. Please ensure you have the necessary permissions."
    exit
}

# Count the number of GPOs
$totalGpos = $gpos.Count

# Display the total number of GPOs
Write-Output "Total number of GPOs: $totalGpos"

# Show GPOs in a grid view
$gpos | Select-Object DisplayName, Id, CreationTime, ModificationTime


# Define the Excel file path
# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$gpos | Export-Excel -Path $excelPath -WorksheetName "AllGPOs" -AutoSize -TableName "AllGPOs" -TableStyle Medium9 -BoldTopRow -FreezeTopRow

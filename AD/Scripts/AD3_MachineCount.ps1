# Import the Active Directory module
Import-Module ActiveDirectory

# Get all computer objects from Active Directory that are not disabled
$computers = Get-ADComputer -Filter {Enabled -eq $true} -Property Name

# Create a custom object to hold the total count
$totalCount = [PSCustomObject]@{
    'Total Active Computers' = $computers.Count
}


# Define the Excel file path
# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$totalCount | Export-Excel -Path $excelPath -WorksheetName "MachineCount" -AutoSize -TableName "MachineCount" -TableStyle Medium9 -BoldTopRow -FreezeTopRow

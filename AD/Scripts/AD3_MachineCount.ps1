# Import the Active Directory module
Import-Module ActiveDirectory

# Calculate the date 60 days ago
$60DaysAgo = (Get-Date).AddDays(-60).ToFileTime()

# Get all computer objects from Active Directory that are not disabled and have logged on in the past 60 days
$computers = Get-ADComputer -Filter {Enabled -eq $true -and lastLogonTimestamp -ge $60DaysAgo} -Properties Name, lastLogonTimestamp

# Create a custom object to hold the total count
$totalCount = [PSCustomObject]@{
    'Total Active Computers (Last 60 Days)' = $computers.Count
}


# Define the Excel file path
# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$totalCount | Export-Excel -Path $excelPath -WorksheetName "MachineCount" -AutoSize -TableName "MachineCount" -TableStyle Medium9 -BoldTopRow -FreezeTopRow

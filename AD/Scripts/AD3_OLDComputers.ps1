# Import the Active Directory module
Import-Module ActiveDirectory

# Set the number of days
$days = 180

# Calculate the date for comparison
$timeSpan = (Get-Date).AddDays(-$days)

# Get all computer accounts that haven't logged in for more than the specified days
$staleComputers = Get-ADComputer -Filter {LastLogonDate -lt $timeSpan} -Properties LastLogonDate |
    Where-Object { $_.LastLogonDate -ne $null }

# Display the results in an Out-GridView
#$staleComputers | Select-Object Name, LastLogonDate | Out-GridView -Title "Stale Workstations (Not Logged In Over 180 Days)"


# Define the Excel file path
# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$staleComputers | Export-Excel -Path $excelPath -WorksheetName "StaleComputers" -AutoSize -TableName "StaleComputers" -TableStyle Medium9 -BoldTopRow -FreezeTopRow

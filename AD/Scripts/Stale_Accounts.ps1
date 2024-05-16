Import-Module ActiveDirectory
Import-Module ImportExcel

# Define the threshold for inactivity (e.g., 90 days)
$thresholdDate = (Get-Date).AddDays(-90)

# Get stale user accounts
$staleUsers = Get-ADUser -Filter {LastLogonDate -lt $thresholdDate -and Enabled -eq $true} -Properties LastLogonDate | Select-Object Name, SamAccountName, LastLogonDate

# Get stale computer accounts
$staleComputers = Get-ADComputer -Filter {LastLogonDate -lt $thresholdDate -and Enabled -eq $true} -Properties LastLogonDate | Select-Object Name, SamAccountName, LastLogonDate

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$staleUsers | Export-Excel -Path $excelPath -WorksheetName "Stale Users" -AutoSize -TableName "StaleUsers" -TableStyle Medium18 -Append
$staleComputers | Export-Excel -Path $excelPath -WorksheetName "Stale Computers" -AutoSize -TableName "StaleComputers" -TableStyle Medium19 -Append

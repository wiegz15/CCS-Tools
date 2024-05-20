Import-Module ActiveDirectory
Import-Module ImportExcel

# Define the threshold for inactivity (e.g., 90 days)
$thresholdDate = (Get-Date).AddDays(-90)

# Get stale server accounts
$staleServers = Get-ADComputer -Filter {LastLogonDate -lt $thresholdDate -and Enabled -eq $true -and OperatingSystem -like '*Server*'} -Properties LastLogonDate | Select-Object Name, SamAccountName, LastLogonDate

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$staleServers | Export-Excel -Path $excelPath -WorksheetName "Stale Servers" -AutoSize -TableName "StaleServers" -TableStyle Medium19 -Append

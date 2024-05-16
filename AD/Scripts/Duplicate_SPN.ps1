Import-Module ActiveDirectory
Import-Module ImportExcel

# Get all user accounts with SPNs
$userSPNs = Get-ADUser -Filter {ServicePrincipalName -ne $null} -Properties ServicePrincipalName | Select-Object Name, SamAccountName, ServicePrincipalName

# Get all computer accounts with SPNs
$computerSPNs = Get-ADComputer -Filter {ServicePrincipalName -ne $null} -Properties ServicePrincipalName | Select-Object Name, SamAccountName, ServicePrincipalName

# Combine results into a single array
$allSPNs = $userSPNs + $computerSPNs

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$allSPNs | Export-Excel -Path $excelPath -WorksheetName "SPNs" -AutoSize -TableName "SPNs" -TableStyle Medium22 -Append

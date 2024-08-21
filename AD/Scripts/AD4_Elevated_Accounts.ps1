Import-Module ActiveDirectory
Import-Module ImportExcel

# Define groups to check for elevated privileges
$elevatedGroups = @("Domain Admins", "Enterprise Admins", "Schema Admins", "Administrators")

$elevatedAccounts = @()

# Loop through each elevated group
foreach ($group in $elevatedGroups) {
    $members = Get-ADGroupMember -Identity $group -Recursive | Where-Object { $_.objectClass -eq 'user' }
    foreach ($member in $members) {
        $lastLogon = (Get-ADUser -Identity $member.SamAccountName -Properties LastLogonDate).LastLogonDate

        $elevatedAccounts += [PSCustomObject]@{
            GroupName   = $group
            AccountName = $member.SamAccountName
            LastLogon   = $lastLogon
        }
    }
}

# Count the number of unique elevated accounts
$uniqueElevatedAccountsCount = ($elevatedAccounts | Select-Object -Property AccountName -Unique).Count

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"

# Create a summary sheet with the count of unique elevated accounts
$summary = @([PSCustomObject]@{ ElevatedAccountsCount = $uniqueElevatedAccountsCount })
$summary | Export-Excel -Path $excelPath -WorksheetName "Elevated Privileges Summary" -AutoSize -TableName "ElevatedAccountsSummary" -TableStyle Medium16 -Append

# Export the detailed list of elevated accounts and their last login times
$elevatedAccounts | Export-Excel -Path $excelPath -WorksheetName "Elevated Privileges" -AutoSize -TableName "ElevatedAccounts" -TableStyle Medium17 -Append

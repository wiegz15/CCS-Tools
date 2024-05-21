# Import the necessary modules
Import-Module ActiveDirectory
Import-Module ImportExcel

# Get all user accounts with SPNs
$userSPNs = Get-ADUser -Filter * -Properties ServicePrincipalName | Where-Object { $_.ServicePrincipalName -ne $null } | Select-Object Name, SamAccountName, ServicePrincipalName

# Get all computer accounts with SPNs
$computerSPNs = Get-ADComputer -Filter * -Properties ServicePrincipalName | Where-Object { $_.ServicePrincipalName -ne $null } | Select-Object Name, SamAccountName, ServicePrincipalName

# Combine results into a single array
$allSPNs = $userSPNs + $computerSPNs

# Flatten the SPNs into a single list with account information
$spnList = @()
foreach ($account in $allSPNs) {
    foreach ($spn in $account.ServicePrincipalName) {
        $spnList += [PSCustomObject]@{
            AccountName = $account.SamAccountName
            SPN         = $spn
        }
    }
}

# Find duplicate SPNs
$duplicateSPNs = $spnList | Group-Object SPN | Where-Object { $_.Count -gt 1 }

# Prepare output
$output = @()
if ($duplicateSPNs.Count -eq 0) {
    $output += [PSCustomObject]@{
        AccountName = "None"
        SPN         = "No duplicate SPNs found"
    }
} else {
    foreach ($group in $duplicateSPNs) {
        foreach ($item in $group.Group) {
            $output += [PSCustomObject]@{
                AccountName = $item.AccountName
                SPN         = $item.SPN
            }
        }
    }
}

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_DuplicateSPN.xlsx"
$output | Export-Excel -Path $excelPath -WorksheetName "Duplicate SPNs" -AutoSize -TableName "DuplicateSPNs" -TableStyle Medium22

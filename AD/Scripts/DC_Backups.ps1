Import-Module ActiveDirectory
Import-Module ImportExcel

# Get the domain DN dynamically
$domainDN = (Get-ADDomain).DistinguishedName

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *

$backups = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    try {
        $backupStatus = Get-ADObject -Identity "CN=System,$domainDN" -Properties "whenChanged" -Server $dc.HostName
        $backups += [PSCustomObject]@{
            DCName           = $dc.HostName
            LatestBackupDate = $backupStatus.whenChanged
        }
    } catch {
        $backups += [PSCustomObject]@{
            DCName           = $dc.HostName
            LatestBackupDate = "Failed to retrieve"
        }
    }
}

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$backups | Export-Excel -Path $excelPath -WorksheetName "Backups" -AutoSize -TableName "BackupsInfo" -TableStyle Medium11 -Append

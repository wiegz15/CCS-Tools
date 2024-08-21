Import-Module ActiveDirectory
Import-Module ImportExcel

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *

$replicationTimes = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    $dcName = $dc.HostName
    $dcSiteName = $dc.Site

    # Get Replication Times
    $replicationStatus = Get-ADReplicationPartnerMetadata -Target $dcName -Scope Server
    foreach ($status in $replicationStatus) {
        $replicationTimes += [PSCustomObject]@{
            DCName                  = $dcName
            SiteName                = $dcSiteName
            Partner                 = $status.PartnerServer
            LastReplicationAttempt  = $status.LastReplicationAttempt
            LastReplicationSuccess  = $status.LastReplicationSuccess
            LastReplicationResult   = $status.LastReplicationResult
        }
    }
}

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$replicationTimes | Export-Excel -Path $excelPath -WorksheetName "Sites-ReplicationTimes" -AutoSize -TableName "Sites-ReplicationTimes" -TableStyle Medium14 -Append

# Import necessary modules
Import-Module ActiveDirectory
Import-Module ImportExcel

# Get the domain DN dynamically
$domainDN = (Get-ADDomain).DistinguishedName

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *

$timeSyncInfo = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    # Check time synchronization status
    $timeSyncStatus = w32tm /query /status /computer:$dc.HostName | Select-String "Source"

    $timeSyncInfo += [PSCustomObject]@{
        Domain           = $domainDN
        NameOfDC         = $dc.HostName
        IPV4Address      = $dc.IPv4Address
        TimeSyncSource   = if ($timeSyncStatus) {($timeSyncStatus -split ":")[1].Trim()} else {"Unknown"}
    }
}


# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$timeSyncInfo | Export-Excel -Path $excelPath -WorksheetName "Time Sync Status" -AutoSize -TableName "TimeSyncInfo" -TableStyle Medium9 -BoldTopRow -FreezeTopRow -Append


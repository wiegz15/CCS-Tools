# Import necessary modules
Import-Module ActiveDirectory
Import-Module DNSServer
Import-Module ImportExcel

# Get the domain DN dynamically
$domainDN = (Get-ADDomain).DistinguishedName

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *
$generalInfo = @()
$dnsSettings = @()
$backups = @()
$printStatus = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $dc.HostName
    $smbv1Status = (Get-WindowsFeature FS-SMB1 -ComputerName $dc.HostName).Installed
    $spoolerService = Get-Service -Name Spooler -ComputerName $dc.HostName
    $dnsZones = Get-DnsServerZone -ComputerName $dc.HostName
    $zonesWithScavenging = $dnsZones | Where-Object { $_.Aging -eq $true }
    $dnsForwarders = Get-DnsServerForwarder -ComputerName $dc.HostName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress | ForEach-Object { $_.IPAddressToString }    
    $dnsForwarderString = if ($dnsForwarders) { $dnsForwarders -join ', ' } else { "No forwarders configured" }

    # General Info
    $generalInfo += [PSCustomObject]@{
        Domain           = $domainDN
        NameOfDC         = $dc.HostName
        IPV4Address      = $dc.IPv4Address
        SMBv1Enabled     = if ($smbv1Status) {"Enabled"} else {"Disabled"}
        OS               = $osInfo.Caption
        OSBuild          = $osInfo.BuildNumber
        FSMO             = ($dc.OperationMasterRoles -join ', ')
    }

    # DNS Settings    
    $dnsSettings += [PSCustomObject]@{
        DCName           = $dc.HostName
        Scavenging       = $zonesWithScavenging.Count -gt 0
        Forwarders       = $dnsForwarderString
        ZonesScavenging  = $zonesWithScavenging.Count
        RootHints        = (Get-DnsServerRootHint -ComputerName $dc.HostName).Count
    }

    # Backups
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

    # Print Status
    $printStatus += [PSCustomObject]@{
        DCName               = $dc.HostName
        PrintSpoolerStatus   = $spoolerService.Status
        PrintSpoolerStartup  = $spoolerService.StartType
    }
}

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$generalInfo | Export-Excel -Path $excelPath -WorksheetName "General Info" -AutoSize -TableName "GeneralInfo" -TableStyle Medium9 -BoldTopRow -FreezeTopRow
$dnsSettings | Export-Excel -Path $excelPath -WorksheetName "DNS Settings" -AutoSize -TableName "DNSSettings" -TableStyle Medium10 -Append
$backups | Export-Excel -Path $excelPath -WorksheetName "Backups" -AutoSize -TableName "BackupsInfo" -TableStyle Medium11 -Append
$printStatus | Export-Excel -Path $excelPath -WorksheetName "Print Status" -AutoSize -TableName "PrintInfo" -TableStyle Medium12 -Append

Import-Module ActiveDirectory
Import-Module DNSServer
Import-Module ImportExcel

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *

$dnsSettings = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    $dnsZones = Get-DnsServerZone -ComputerName $dc.HostName
    $zonesWithScavenging = $dnsZones | Where-Object { $_.Aging -eq $true }
    $dnsForwarders = Get-DnsServerForwarder -ComputerName $dc.HostName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress | ForEach-Object { $_.IPAddressToString }    
    $dnsForwarderString = if ($dnsForwarders) { $dnsForwarders -join ', ' } else { "No forwarders configured" }

    $dnsSettings += [PSCustomObject]@{
        DCName           = $dc.HostName
        Scavenging       = $zonesWithScavenging.Count -gt 0
        Forwarders       = $dnsForwarderString
        ZonesScavenging  = $zonesWithScavenging.Count
        RootHints        = (Get-DnsServerRootHint -ComputerName $dc.HostName).Count
    }
}

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$dnsSettings | Export-Excel -Path $excelPath -WorksheetName "DNS Settings" -AutoSize -TableName "DNSSettings" -TableStyle Medium10 -Append

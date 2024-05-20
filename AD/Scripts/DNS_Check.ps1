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

    # Fetch aging details for each zone
    $zoneAgingDetails = $dnsZones | ForEach-Object {
        Get-DnsServerZoneAging -ComputerName $dc.HostName -Name $_.ZoneName
    } | Select-Object ZoneName, AvailForScavengeTime, @{Name='ScavengingEnabled'; Expression={$_.AgingEnabled}}

    # Combine details in custom object
    foreach ($zoneDetail in $zoneAgingDetails) {
        $dnsSettings += [PSCustomObject]@{
            DCName              = $dc.HostName
            ZoneName            = $zoneDetail.ZoneName
            Scavenging          = $zoneDetail.ScavengingEnabled
            AvailForScavengeTime = $zoneDetail.AvailForScavengeTime
            Forwarders          = $dnsForwarderString
            RootHints           = (Get-DnsServerRootHint -ComputerName $dc.HostName).Count
        }
    }
}

# Optionally, export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$dnsSettings | Export-Excel -Path $excelPath -WorksheetName "DNS Settings" -AutoSize -TableName "DNSSettings" -TableStyle Medium10 -Append

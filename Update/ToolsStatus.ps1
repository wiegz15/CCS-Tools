$featuresToCheck = 'RSAT-Role-Tools', 'RSAT-AD-Tools', 'RSAT-AD-PowerShell', 'RSAT-ADDS', 'RSAT-ADDS-Tools', 'RSAT-ADLDS', 'RSAT-RDS-Tools', 'RSAT-DNS-Server'
$modulesToCheck = 'ImportExcel', 'VMware.PowerCLI', 'GroupPolicy', 'ActiveDirectory', 'DNSServer'

$installedFeatures = Get-WindowsFeature | Where-Object {$_.Installed -eq $true} | Select-Object -ExpandProperty Name
$status = @()

foreach ($feature in $featuresToCheck) {
    if ($installedFeatures -contains $feature) {
        $status += "${feature}: Installed"
    } else {
        $status += "${feature}: Not Installed"
    }
}

foreach ($module in $modulesToCheck) {
    $moduleCheck = Get-Module -ListAvailable -Name $module
    if ($null -ne $moduleCheck) {
        $status += "${module}: Installed"
    } else {
        $status += "${module}: Not Installed"
    }
}

$status | ForEach-Object { Write-Output $_ }

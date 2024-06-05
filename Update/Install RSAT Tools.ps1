$featuresToInstall = 'RSAT-Role-Tools', 'RSAT-AD-Tools', 'RSAT-AD-PowerShell', 'RSAT-ADDS', 'RSAT-ADDS-Tools', 'RSAT-ADLDS', 'RSAT-RDS-Tools', 'RSAT-DNS-Server'

$installedFeatures = Get-WindowsFeature | Where-Object {$_.Installed -eq $true} | Select-Object -ExpandProperty Name

foreach ($feature in $featuresToInstall) {
    if ($installedFeatures -contains $feature) {
        Write-Output "$feature is already installed."
    } else {
        Write-Output "Installing $feature..."
        Install-WindowsFeature -Name $feature -IncludeAllSubFeature -IncludeManagementTools
        if ($?) {
            Write-Output "$feature installed successfully."
        } else {
            Write-Output "Failed to install $feature."
        }
    }
}

# Import necessary modules
Import-Module ActiveDirectory
Import-Module ImportExcel

# Get the domain DN dynamically
$domainDN = (Get-ADDomain).DistinguishedName

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *

# Get the Forest Functional Level
$forest = Get-ADForest
$forestFunctionalLevel = $forest.ForestMode

$generalInfo = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $dc.HostName
    $smbv1Status = (Get-WindowsFeature FS-SMB1 -ComputerName $dc.HostName).Installed

    $generalInfo += [PSCustomObject]@{
        Domain               = $domainDN
        NameOfDC             = $dc.HostName
        IPV4Address          = $dc.IPv4Address
        OS                   = $osInfo.Caption
        OSBuild              = $osInfo.BuildNumber
        FSMO                 = ($dc.OperationMasterRoles -join ', ')
        ForestFunctionalLevel = $forestFunctionalLevel
    }
}

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$generalInfo | Export-Excel -Path $excelPath -WorksheetName "General Info" -AutoSize -TableName "GeneralInfo" -TableStyle Medium9 -BoldTopRow -FreezeTopRow

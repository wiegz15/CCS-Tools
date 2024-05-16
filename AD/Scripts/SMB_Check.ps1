Import-Module ActiveDirectory
Import-Module ImportExcel

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *

$smbVersions = @()

# Function to check SMB version status
function Get-SMBStatus {
    param (
        [string]$ComputerName,
        [string]$SMBVersion
    )
    $regPath = "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
    $regValue = "SMB$SMBVersion"

    try {
        $smbStatus = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            param ($regPath, $regValue)
            Get-ItemProperty -Path "HKLM:\$regPath" -Name $regValue -ErrorAction Stop
        } -ArgumentList $regPath, $regValue
        return $true
    } catch {
        return $false
    }
}

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    $dcName = $dc.HostName

    $smb1Enabled = Get-SMBStatus -ComputerName $dcName -SMBVersion 1
    $smb2Enabled = Get-SMBStatus -ComputerName $dcName -SMBVersion 2
    $smb3Enabled = Get-SMBStatus -ComputerName $dcName -SMBVersion 3

    $smbVersions += [PSCustomObject]@{
        DCName     = $dcName
        SMB1Enabled = if ($smb1Enabled) {"Enabled"} else {"Disabled"}
        SMB2Enabled = if ($smb2Enabled) {"Enabled"} else {"Disabled"}
        SMB3Enabled = if ($smb3Enabled) {"Enabled"} else {"Disabled"}
    }
}

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$smbVersions | Export-Excel -Path $excelPath -WorksheetName "SMB Versions" -AutoSize -TableName "SMBVersions" -TableStyle Medium15 -Append

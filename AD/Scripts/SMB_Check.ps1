# Import the Active Directory module
Import-Module ActiveDirectory

# Get all domain controllers in the domain
$domainControllers = Get-ADDomainController -Filter *

# Initialize an array to hold the results
$results = @()

# Loop through each domain controller
foreach ($dc in $domainControllers) {
    # Use Invoke-Command to run the Get-SmbServerConfiguration cmdlet on each domain controller
    $smbConfig = Invoke-Command -ComputerName $dc.HostName -ScriptBlock {
        Get-SmbServerConfiguration | Select-Object PSComputerName, EnableSMB1Protocol, EnableSMB2Protocol
    }
    
    # Add the result to the results array
    $results += $smbConfig
}

# Note: This script requires appropriate permissions to run commands remotely and access AD information.
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$results | Export-Excel -Path $excelPath -WorksheetName "SMB Versions" -AutoSize -TableName "SMBVersions" -TableStyle Medium15 -Append
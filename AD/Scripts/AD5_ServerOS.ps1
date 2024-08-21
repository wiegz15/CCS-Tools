# Ensure the Active Directory module is available
Import-Module ActiveDirectory

# Get all computer objects with the Server role
$servers = Get-ADComputer -Filter 'OperatingSystem -like "*Server*"' -Property Name, OperatingSystem

# Select relevant properties
$serverInfo = $servers | Select-Object Name, OperatingSystem


# Note: This script requires appropriate permissions to run commands remotely and access AD information.
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$serverInfo | Export-Excel -Path $excelPath -WorksheetName "ServerOS" -AutoSize -TableName "ServerOS" -TableStyle Medium15 -Append
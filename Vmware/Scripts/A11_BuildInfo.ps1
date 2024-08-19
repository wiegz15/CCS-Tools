# Retrieve all hosts and their ESXi versions
$hostData = Get-VMHost | Select-Object Name, Version, Build

# Display the information in Out-GridView
# $hostData | Out-GridView -Title "ESXi Host Versions"


$worksheetName = "ESXHostBuildInfo"

# Export the results to an Excel file
$hostData | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName $WorksheetName



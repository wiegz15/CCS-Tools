$esx_all = Get-VMHost | Get-View
$Report = @()
foreach ($esx in $esx_all) {
    foreach ($triggered in $esx.TriggeredAlarmState) {
        If ($triggered.OverallStatus -like "red") {
            $lineitem = New-Object psobject -Property @{
                Name = $esx.Name
                AlarmInfo = (Get-View -Id $triggered.Alarm).Info.Name
            }
            $Report += $lineitem
        }
    }
}
#$Report | Sort-Object Name | Out-GridView

# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "TriggeredAlarms"

# Export the results to an Excel file
$Report | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "TriggeredAlarms"


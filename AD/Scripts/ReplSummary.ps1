# Import modules
Import-Module ADEssentials
Import-Module ImportExcel

# Get replication summary
$replicationSummary = Get-WinADForestReplicationSummary

# Define file path
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"

# Convert the replication summary to a format suitable for Export-Excel
$replicationSummaryFormatted = $replicationSummary | Select-Object Server, LargestDelta, Fails, Total, PercentageError, Type, ReplicationError

# Export to Excel
$replicationSummaryFormatted | Export-Excel -Path $excelPath -WorksheetName "ReplicationSummary" -AutoSize -TableName "ReplicationSummary" -TableStyle Medium14 -Append


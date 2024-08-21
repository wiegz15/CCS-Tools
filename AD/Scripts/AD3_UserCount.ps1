# Import the Active Directory module
Import-Module ActiveDirectory

# Get all enabled users in Active Directory
$activeUsers = Get-ADUser -Filter {Enabled -eq $true} -Properties Enabled

# Count the total number of active users
$totalActiveUsers = $activeUsers.Count

# Create a custom object to display in Out-GridView
$result = [PSCustomObject]@{
    "Total Active Users" = $totalActiveUsers
}

# Display the result in Out-GridView
# $result | Out-GridView


# Define file path
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"

# Convert the replication summary to a format suitable for Export-Excel
$replicationSummaryFormatted = $replicationSummary | Select-Object Server, LargestDelta, Fails, Total, PercentageError, Type, ReplicationError

# Export to Excel
$result | Export-Excel -Path $excelPath -WorksheetName "activeusers" -AutoSize -TableName "activeusers" -TableStyle Medium14 -Append


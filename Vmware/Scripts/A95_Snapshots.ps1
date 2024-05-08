# Set default snapshot age value to 0
$textAge = '0'

# Find old snapshots and include the VM name
$snapshots = Get-VM | Get-Snapshot | Where-Object { $_.Created -lt (Get-Date).AddDays(-[int]$textAge) } | ForEach-Object {
    # Create a custom PSObject that includes VM Name, Snapshot Name, Created Date, and Size
    [PSCustomObject]@{
        VMName = $_.VM.Name
        SnapshotName = $_.Name
        Created = $_.Created
        SizeMB = $_.SizeMB
    }
}

# Display snapshots in Out-GridView, now including VM names
# $selectedSnapshots = $snapshots | Out-GridView -PassThru -Title "Select Snapshots to Delete"


# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "Snapshots"

# Export the results to an Excel file
$snapshots | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table21"


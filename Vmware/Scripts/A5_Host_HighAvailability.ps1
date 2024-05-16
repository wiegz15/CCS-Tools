
# Retrieve all clusters
$clusters = Get-Cluster

# Initialize an array to hold results
$results = @()

foreach ($cluster in $clusters) {
    # Check the HA status
    $haEnabled = $cluster.HAEnabled

    # Add cluster status and HA status to the results array
    $results += [PSCustomObject]@{
        ClusterName = $cluster.Name
        HAEnabled = if ($haEnabled) { "Yes" } else { "No" }
    }
}

# Display the results in an Out-GridView
# $results | Out-GridView -Title "Cluster Status and HA Enabled"

# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "HA"

# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "HA"




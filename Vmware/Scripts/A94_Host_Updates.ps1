
# Retrieve all hosts
$vmHosts = Get-VMHost

# Prepare an array to hold the compliance results
$complianceResults = @()

foreach ($vmHost in $vmHosts) {
    # Get compliance information for each host
    $complianceInfo = Get-Compliance -Entity $vmHost -Detailed
    
    # Filter for baselines that contain "Critical Host Patches"
    $criticalPatchesCompliance = $complianceInfo | Where-Object { $_.Baseline.Name -like "*Patches*" }
    
    if ($criticalPatchesCompliance) {
        # If a matching baseline is found, prepare the result object
        foreach ($baseline in $criticalPatchesCompliance) {
            $result = [PSCustomObject]@{
                HostName = $vmHost.Name
                BaselineName = $baseline.Baseline.Name
                ComplianceStatus = $baseline.Status
            }

            # Add the result to the results array
            $complianceResults += $result
        }
    } else {
        # If the baseline is not attached to the host
        $result = [PSCustomObject]@{
            HostName = $vmHost.Name
            BaselineName = "Critical Host Patches (Not Attached)"
            ComplianceStatus = "N/A"
        }

        # Add the result to the results array
        $complianceResults += $result
    }
}

# Display the results in a grid view
# $complianceResults | Out-GridView


# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "HostUptime"

# Export the results to an Excel file
$complianceResults | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table22"




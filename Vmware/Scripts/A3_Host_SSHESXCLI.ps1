# Get all host servers
$vmHosts = Get-VMHost

# Prepare an array to hold the results
$results = @()

foreach ($vmHost in $vmHosts) {
    # Retrieve the ESXi Shell and SSH services
    $esxiShellService = Get-VMHostService -VMHost $vmHost | Where-Object {$_.Key -eq "TSM"}
    $sshService = Get-VMHostService -VMHost $vmHost | Where-Object {$_.Key -eq "TSM-SSH"}

    # Add ESXi Shell status to results
    $results += [PSCustomObject]@{
        HostName = $vmHost.Name
        Service = "ESXi Shell"
        Status = if ($esxiShellService.Running) { "Running" } else { "Stopped" }
    }

    # Add SSH status to results
    $results += [PSCustomObject]@{
        HostName = $vmHost.Name
        Service = "SSH"
        Status = if ($sshService.Running) { "Running" } else { "Stopped" }
    }
}

# Display the results in an Out-GridView
#$results | Out-GridView -Title "ESXi Shell and SSH Services Status"

# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name
$worksheetName = "ESXSSH"

# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table7"

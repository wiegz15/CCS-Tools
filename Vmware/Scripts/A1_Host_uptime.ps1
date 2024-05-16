# Get all ESXi hosts in the vCenter
$esxiHosts = Get-VMHost

# Initialize an array to hold the custom objects
$hostDetails = @()

# For each host, create a custom object with desired properties
foreach ($esxiHost in $esxiHosts) {
    $hostName = $esxiHost.Name
    $uptimeSeconds = $esxiHost.ExtensionData.Summary.QuickStats.Uptime
    $uptime = New-TimeSpan -Seconds $uptimeSeconds
    $connectionState = $esxiHost.ConnectionState
    
    $uptimeString = "$($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes"
    
    # Create custom object for each host
    $hostDetail = [PSCustomObject]@{
        HostName = $hostName
        ConnectionState = $connectionState
        Uptime = $uptimeString
    }
    
    # Add the custom object to the array
    $hostDetails += $hostDetail
}

# Display the results in Out-GridView
#$hostDetails | Out-GridView -Title "ESXi Hosts Status"


# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "HostUptime"

# Export the results to an Excel file
$hostdetails | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "HostUptime"



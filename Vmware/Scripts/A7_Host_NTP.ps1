
# Get all host servers
$vmHosts = Get-VMHost

# Prepare an array to hold the results
$results = @()

foreach ($vmHost in $vmHosts) {
    # Retrieve the NTP service status
    $ntpService = Get-VMHostService -VMHost $vmHost | Where-Object {$_.Key -eq "ntpd"}
    
    # Retrieve NTP server settings
    $ntpSettings = Get-VMHostNtpServer -VMHost $vmHost
    
    # Prepare the result object
    $result = [PSCustomObject]@{
        HostName = $vmHost.Name
        NTPServiceRunning = $ntpService.Running
        NTPServers = ($ntpSettings -join ", ")
    }
    
    # Add the result to the results array
    $results += $result
}

# Display the results in an Out-GridView
#$results | Out-GridView

# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "NTPServer"

# Export the results to an Excel file
$results | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table5"


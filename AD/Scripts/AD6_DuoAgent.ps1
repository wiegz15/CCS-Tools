# Import Active Directory module
Import-Module ActiveDirectory

# Get a list of servers from Active Directory
$servers = Get-ADComputer -Filter {OperatingSystem -like '*Server*'} -Property Name, Enabled | Select-Object Name, Enabled

# Create a results array to store server information
$results = @()

# Total servers for progress tracking
$totalServers = $servers.Count
$currentServer = 0

# Loop through each server to check its status and Duo installation
foreach ($server in $servers) {
    $currentServer++
    $serverName = $server.Name
    $serverStatus = if ($server.Enabled) { "Active" } else { "Disabled" }
    
    # Initialize Duo check to "Not Applicable"
    $duoInstalled = "Not Applicable"

    # Only check for Duo if the server is Active
    if ($serverStatus -eq "Active") {
        $filePath = "\\$serverName\C$\Program Files\Duo Security\WindowsLogon\DuoCredFilter.dll"

        try {
            if (Test-Path $filePath) {
                $duoInstalled = "Yes"
            } else {
                $duoInstalled = "No"
            }
        } catch {
            $duoInstalled = "Unable to Connect"
        }
    }

    # Add server details to the results array
    $results += [PSCustomObject]@{
        ServerName   = $serverName
        Status       = $serverStatus
        DuoInstalled = $duoInstalled
    }

    # Display progress in the console
    Write-Host "[$currentServer/$totalServers] Server: $serverName - Status: $serverStatus - Duo Installed: $duoInstalled"
}


# Output the results to an Out-GridView
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$results | Export-Excel -Path $excelPath -WorksheetName "DuoAgent" -AutoSize -TableName "DuoAgent" -TableStyle Medium12 -Append

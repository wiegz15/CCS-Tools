# Import the Active Directory module
Import-Module ActiveDirectory

# Get all server objects from Active Directory
$servers = Get-ADComputer -Filter {OperatingSystem -Like "*Server*"} -Property Name

# Prepare an array to store the results
$results = @()

foreach ($server in $servers) {
    # Output the name of the server being checked
    Write-Host "Checking server: $($server.Name)"

    # Initialize the result object
    $result = New-Object -TypeName PSObject
    $result | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $server.Name

    # Attempting DNS check and retrieving IP address
    Write-Host "Attempting DNS check for $($server.Name)..."
    try {
        $dnsResult = Resolve-DnsName $server.Name -ErrorAction Stop
        $result | Add-Member -MemberType NoteProperty -Name "DNS" -Value "Found"
        $result | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $dnsResult.IPAddress
        Write-Host "DNS check successful for $($server.Name), IP: $($dnsResult.IPAddress)"
    } catch {
        $result | Add-Member -MemberType NoteProperty -Name "DNS" -Value "Not Found"
        $result | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value "N/A"
        Write-Host "DNS check failed for $($server.Name)"
    }

    # Output connection test status
    Write-Host "Attempting Test-Connection for $($server.Name)..."
    $connectionTest = Test-Connection $server.Name -Count 1 -Quiet
    if ($connectionTest) {
        $result | Add-Member -MemberType NoteProperty -Name "Online" -Value "Yes"
        Write-Host "Server $($server.Name) is online."
    } else {
        $result | Add-Member -MemberType NoteProperty -Name "Online" -Value "No"
        Write-Host "Server $($server.Name) is not online."
    }

    # Add result to results array
    $results += $result
}


# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file on the user's desktop
$desktopPath = [System.Environment]::GetFolderPath("Desktop")
$excelPath = Join-Path -Path $desktopPath -ChildPath "Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "VM_Online"

# Output results to GridView
$results | Export-Excel -Path $excelPath -WorksheetName "OrphanedServers" -AutoSize -TableName "Table31"
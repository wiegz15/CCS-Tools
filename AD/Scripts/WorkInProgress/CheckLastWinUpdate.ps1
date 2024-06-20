# Ensure the necessary modules are available
Import-Module ActiveDirectory
Import-Module ImportExcel

# Function to get all Windows Server computers in the domain
function Get-WindowsServerComputers {
    Get-ADComputer -Filter {OperatingSystem -Like "*Windows Server*"} -Property OperatingSystem | 
    Select-Object -ExpandProperty Name
}

# Function to get the last installed update on a server
function Get-LastInstalledUpdate {
    param (
        [string]$ComputerName
    )

    # Check if the server is reachable
    if (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
        try {
            # Get the last installed update
            $lastUpdate = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                Get-WmiObject -Query "SELECT * FROM Win32_QuickFixEngineering" | 
                Sort-Object -Property InstalledOn -Descending |
                Select-Object -First 1 -Property Description, HotFixID, InstalledOn
            }
            
            # Display the result
            [PSCustomObject]@{
                ServerName = $ComputerName
                Description = $lastUpdate.Description
                HotFixID    = $lastUpdate.HotFixID
                InstalledOn = $lastUpdate.InstalledOn
            }
        } catch {
            Write-Host "Failed to retrieve updates from $ComputerName" -ForegroundColor Red
        }
    } else {
        Write-Host "$ComputerName is not reachable." -ForegroundColor Yellow
    }
}

# Get the list of Windows Server computers from the domain
$servers = Get-WindowsServerComputers

# Initialize an array to store results
$results = @()

# Loop through each server and get the last installed update
foreach ($server in $servers) {
    $result = Get-LastInstalledUpdate -ComputerName $server
    if ($result) {
        $results += $result
    }
}

# Convert the results to a DataTable
$resultsTable = $results | ConvertTo-DataTable

# Define the Excel file path
$excelPath = Join-Path -Path $reportsDir -ChildPath "LastWinUpdate.xlsx"

# Export the results to an Excel file with a table
$resultsTable | Export-Excel -Path $excelPath -AutoSize -TableName "LastInstalledUpdates"

Write-Host "Results exported to $excelPath"
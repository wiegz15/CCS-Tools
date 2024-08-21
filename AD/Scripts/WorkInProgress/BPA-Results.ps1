# Import necessary modules
Import-Module ServerManager
Import-Module ActiveDirectory

# Ensure BPA modules are updated
Write-Output "Updating BPA modules..."
Update-Help -Module ServerManager

# Define BPA Model IDs to run
$BpaModels = @(
    "Microsoft/Windows/DirectoryServices",
    "Microsoft/Windows/DNSServer",
    "Microsoft/Windows/FileServices"
)

# Get the current domain
$currentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()

# Get a list of domain controllers in the Domain Controllers OU
$domainControllers = Get-ADComputer -Filter { Name -like "*" } -SearchBase "OU=Domain Controllers,DC=$($currentDomain.Name -replace '\.', ',DC=')" -Properties Name | Select-Object -ExpandProperty Name

# Create arrays to hold results
$allResults = @()

# Function to run BPA checks on a remote server and add the server name to each result
function Get-BPAResultsRemote {
    param (
        [string]$server,
        [array]$BpaModels
    )
    
    Invoke-Command -ComputerName $server -ScriptBlock {
        param ($BpaModels, $serverName)
        $results = @()
        foreach ($model in $BpaModels) {
            Invoke-BPAModel -ModelId $model
            $modelResults = Get-BPAResult -ModelId $model
            foreach ($result in $modelResults) {
                $result | Add-Member -MemberType NoteProperty -Name DomainController -Value $serverName
                $results += $result
            }
        }
        return $results
    } -ArgumentList $BpaModels, $server
}

# Run BPA checks on each domain controller and collect results
foreach ($dc in $domainControllers) {
    Write-Output "Running BPA checks on $dc..."
    $results = Get-BPAResultsRemote -server $dc -BpaModels $BpaModels
    $allResults += $results
}

# Filter out 'Information' severity results
$filteredResults = $allResults | Where-Object { $_.Severity -ne 'Information' }

# Set file paths to the temp directory
$tempPath = [System.IO.Path]::GetTempPath()
$csvFilePath = Join-Path -Path $tempPath -ChildPath "BPAResults.csv"
$htmlFilePath = Join-Path -Path $tempPath -ChildPath "BPAResults.html"

# Export results to CSV
$filteredResults | Export-Csv -Path $csvFilePath -NoTypeInformation

# Export results to HTML
$filteredResults | ConvertTo-Html -Property DomainController, ModelId, Title, Severity, Problem, Impact, Resolution -Title "BPA Results" | Out-File -FilePath $htmlFilePath

# Output file paths for user reference
Write-Output "CSV file saved to: $csvFilePath"
Write-Output "HTML file saved to: $htmlFilePath"

# Open the HTML file
Start-Process $htmlFilePath

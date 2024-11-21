# Import necessary modules
Import-Module ServerManager
Import-Module ActiveDirectory
Import-Module ImportExcel

# Function to get servers with a specific role from Active Directory
function Get-ServersWithRole {
    param (
        [string]$RoleName
    )
    
    switch ($RoleName) {
        "AD-Domain-Services" { $filter = '(ObjectClass -eq "computer") -and (ServicePrincipalName -like "ldap*")' }
        "DNS" { $filter = '(ObjectClass -eq "computer") -and (ServicePrincipalName -like "dns*")' }
        default { Write-Error "RoleName not recognized"; return @() }
    }
    
    $servers = Get-ADComputer -Filter $filter -Property Name | Select-Object -ExpandProperty Name
    return $servers
}

# Get servers for DirectoryServices and DNSServer roles
$DirectoryServicesServers = Get-ServersWithRole -RoleName "AD-Domain-Services"
$DNSServers = Get-ServersWithRole -RoleName "DNS"

# Combine the lists of servers and remove duplicates
$allServers = $DirectoryServicesServers + $DNSServers | Sort-Object -Unique

# Display the list of servers and allow the user to select
$selectedServers = $allServers | Out-GridView -Title "Select Servers to Run BPA Models" -PassThru

# Create arrays to hold results
$directoryServicesResults = @()
$dnsServerResults = @()

# Function to run BPA models and collect results
function Run-BPAModels {
    param (
        [string]$Model,
        [string]$Server
    )
    
    Invoke-Command -ComputerName $Server -ScriptBlock {
        param ($model)
        
        # Run the BPA model and collect results
        try {
            Invoke-BPAModel -ModelId $model
            $modelResults = Get-BPAResult -ModelId $model
            return $modelResults
        } catch {
            Write-Warning "Failed to invoke BPA model $model on server $using:Server"
            return @()
        }
    } -ArgumentList $Model
}

# Run BPA models on selected servers
foreach ($server in $selectedServers) {
    Write-Output "Running DirectoryServices BPA model on server: $server"
    $results = Run-BPAModels -Model "Microsoft/Windows/DirectoryServices" -Server $server
    $directoryServicesResults += $results | ForEach-Object {
        [PSCustomObject]@{
            Server        = $server
            ModelId       = $_.ModelId
            SourceId      = $_.SourceId
            RuleId        = $_.RuleId
            Title         = $_.Title
            Severity      = $_.Severity
            Category      = $_.Category
            Problem       = $_.Problem
            Impact        = $_.Impact
            Resolution    = $_.Resolution
            LastScanTime  = $_.LastScanTime
        }
    }

    Write-Output "Running DNSServer BPA model on server: $server"
    $results = Run-BPAModels -Model "Microsoft/Windows/DNSServer" -Server $server
    $dnsServerResults += $results | ForEach-Object {
        [PSCustomObject]@{
            Server        = $server
            ModelId       = $_.ModelId
            SourceId      = $_.SourceId
            RuleId        = $_.RuleId
            Title         = $_.Title
            Severity      = $_.Severity
            Category      = $_.Category
            Problem       = $_.Problem
            Impact        = $_.Impact
            Resolution    = $_.Resolution
            LastScanTime  = $_.LastScanTime
        }
    }
}

# Filter out Information severity
$directoryServicesFiltered = $directoryServicesResults | Where-Object { $_.Severity -ne 'Information' }
$dnsServerFiltered = $dnsServerResults | Where-Object { $_.Severity -ne 'Information' }

# Export results to an Excel file with two sheets
$excelPath = Join-Path -Path $reportsDir -ChildPath "BPAResults.xlsx"
$directoryServicesFiltered | Export-Excel -Path $excelPath -AutoSize -Title "Directory Services BPA Results" -WorksheetName "DirectoryServices"
$dnsServerFiltered | Export-Excel -Path $excelPath -AutoSize -Title "DNS Server BPA Results" -WorksheetName "DNSServer" -Append


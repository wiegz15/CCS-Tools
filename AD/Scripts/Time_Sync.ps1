#Import Modules
Import-Module ImportExcel
Import-Module ActiveDirectory

# Get all domain controllers in the domain
$domainControllers = Get-ADDomainController -Filter *

# Prepare a variable to store the results
$results = @()

# Script block to check NTP settings and collect data
$scriptBlock = {
    $dcName = $env:COMPUTERNAME
    $timeSource = w32tm /query /source
    $timeStatus = w32tm /query /status

    # Return an object with properties
    return [PSCustomObject]@{
        ComputerName = $dcName
        TimeSource = $timeSource
        TimeStatus = $timeStatus -join "`n"  # Join all lines to a single string
    }
}

# Run the script block on each domain controller and collect results
foreach ($dc in $domainControllers) {
    $result = Invoke-Command -ComputerName $dc.HostName -ScriptBlock $scriptBlock
    $results += $result
}


# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$results | Select-Object ComputerName, TimeSource, TimeStatus | Export-Excel -Path $excelPath -WorksheetName "Time Sync Status" -AutoSize -TableName "TimeSyncInfo" -TableStyle Medium9 -BoldTopRow -FreezeTopRow -Append


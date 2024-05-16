Import-Module ActiveDirectory
Import-Module ImportExcel

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *

$printStatus = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    $spoolerService = Get-Service -Name Spooler -ComputerName $dc.HostName

    $printStatus += [PSCustomObject]@{
        DCName               = $dc.HostName
        PrintSpoolerStatus   = $spoolerService.Status
        PrintSpoolerStartup  = $spoolerService.StartType
    }
}

# Export to Excel
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$printStatus | Export-Excel -Path $excelPath -WorksheetName "Print Status" -AutoSize -TableName "PrintInfo" -TableStyle Medium12 -Append

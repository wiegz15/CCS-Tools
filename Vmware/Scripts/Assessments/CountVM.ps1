Import-Module ImportExcel

# Get the count of powered-on VMs
$poweredOnVMCount = (Get-VM | Where-Object { $_.PowerState -eq "PoweredOn" }).Count

# Output the count
Write-Output "Number of powered-on VMs: $poweredOnVMCount"


# Export the cluster data to the first sheet in the Excel file
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMCount.xlsx"
$worksheetName = "VM Count"
$poweredOnVMCount | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "RunningVMCount"


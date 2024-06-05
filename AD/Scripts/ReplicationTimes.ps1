$myRepInfo = @(repadmin /replsum * /bysrc /bydest /sort:delta)
 
# Initialize our array.
$cleanRepInfo = @()
# Start @ #10 because all the previous lines are junk formatting
# and strip off the last 4 lines because they are not needed.
for ($i=10; $i -lt ($myRepInfo.Count-4); $i++) {
    if ($myRepInfo[$i] -ne "") {
        # Remove empty lines from our array.
        $myRepInfo[$i] -replace '\s+', " "
        $cleanRepInfo += $myRepInfo[$i]
    }
}

$finalRepInfo = @()
foreach ($line in $cleanRepInfo) {
    $splitRepInfo = $line -split '\s+', 8
    if ($splitRepInfo[0] -eq "Source") { $repType = "Source" }
    if ($splitRepInfo[0] -eq "Destination") { $repType = "Destination" }

    if ($splitRepInfo[1] -notmatch "DSA") {
        # Create an Object and populate it with our values.
        $objRepValues = New-Object PSObject -Property @{
            DSAType  = $repType      # Source or Destination DSA
            Hostname = $splitRepInfo[1] # Hostname
            Delta    = $splitRepInfo[2] # Largest Delta
            Fails    = $splitRepInfo[3] # Failures
            Total    = $splitRepInfo[5] # Totals
            'Error%' = $splitRepInfo[6] # % errors  
            ErrorMsg = $splitRepInfo[7] # Error code
        }

        # Add the Object as a row to our array   
        $finalRepInfo += $objRepValues
    }
}

# Convert array to a DataTable
# $dataTable = ConvertTo-DataTable -array $finalRepInfo

# Define the Excel file path
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"

# Export DataTable to Excel using ImportExcel module
$finalRepInfo | Export-Excel -Path $excelPath -WorkSheetname "ReplicationReport" -AutoSize -TableName "ReplicationReport" -TableStyle Medium14 -Append

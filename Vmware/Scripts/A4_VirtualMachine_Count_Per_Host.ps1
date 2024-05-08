# Retrieve all hosts and count VMs on each
$vmCountPerHost = Get-VMHost | ForEach-Object {
    $hostName = $_.Name
    $vmCount = (Get-VM -Location $_).Count
    [PSCustomObject]@{
        HostName = $hostName
        VMCount = $vmCount
    }
}

# Display the results in a grid view
#$vmCountPerHost | Out-GridView -Title "VM Count per Host"



# Import the ImportExcel module
Import-Module ImportExcel

# Define the path for the Excel file in the Reports directory
$excelPath = Join-Path -Path $reportsDir -ChildPath "VMware_Output.xlsx"

# Get the name of the current script for the worksheet name

$worksheetName = "vmCountPerHost"

# Export the results to an Excel file
$vmCountPerHost | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName "Table9"


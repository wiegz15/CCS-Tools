# Import the Active Directory module
Import-Module ActiveDirectory

# Get all OUs
$OUs = Get-ADOrganizationalUnit -Filter * | Sort-Object Name

# Create an array to store the results
$Results = @()

# Loop through each OU
foreach ($OU in $OUs) {
    # Get the number of computers in the current OU
    $ComputerCount = (Get-ADComputer -Filter * -SearchBase $OU.DistinguishedName).Count

    # Create a custom object with the OU name and computer count
    $Result = [PSCustomObject]@{
        'Organizational Unit' = $OU.Name
        'Computer Count' = $ComputerCount
    }

    # Add the result to the array
    $Results += $Result
}


# Note: This script requires appropriate permissions to run commands remotely and access AD information.
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$results | Export-Excel -Path $excelPath -WorksheetName "OUsComputers" -AutoSize -TableName "OUsComputers" -TableStyle Medium15 -Append
# Import the Active Directory module
Import-Module ActiveDirectory

# Get all OUs
$OUs = Get-ADOrganizationalUnit -Filter * | Sort-Object Name

# Create an array to store the results
$Results = @()

# Loop through each OU
foreach ($OU in $OUs) {
    # Get the number of users in the current OU
    $UserCount = (Get-ADUser -Filter * -SearchBase $OU.DistinguishedName).Count

    # Create a custom object with the OU name and user count
    $Result = [PSCustomObject]@{
        'Organizational Unit' = $OU.Name
        'User Count' = $UserCount
    }

    # Add the result to the array
    $Results += $Result
}

# Note: This script requires appropriate permissions to run commands remotely and access AD information.
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$results | Export-Excel -Path $excelPath -WorksheetName "OUsUsers" -AutoSize -TableName "OUsUsers" -TableStyle Medium15 -Append
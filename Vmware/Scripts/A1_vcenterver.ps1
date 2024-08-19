# Get vCenter version
$vcVersion = $global:DefaultVIServer.ExtensionData.Content.About

# Create custom object with version information
$versionInfo = [PSCustomObject]@{
    'vCenter Name' = $global:DefaultVIServer.Name
    'Version' = $vcVersion.Version
    'Build' = $vcVersion.Build
    'Full Name' = $vcVersion.FullName
    'OS Type' = $vcVersion.OsType
    'Product Line' = $vcVersion.ProductLineId
    'API Type' = $vcVersion.ApiType
    'API Version' = $vcVersion.ApiVersion
}

# Display in Out-GridView
# $versionInfo | Out-GridView -Title "vCenter Version Information"
$worksheetName = "VcenterInfo"

# Export the results to an Excel file
$versionInfo | Export-Excel -Path $excelPath -WorksheetName $worksheetName -AutoSize -TableName $WorksheetName



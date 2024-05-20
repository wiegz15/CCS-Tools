# Import GroupPolicy module if not already loaded
if (-not (Get-Module -Name GroupPolicy)) {
    Import-Module GroupPolicy
}

# Get all GPOs in the domain
$allGPOs = Get-GPO -All

# Filter out GPOs that are not linked to any OU
$unlinkedGPOs = foreach ($gpo in $allGPOs) {
    $reportXml = [xml](Get-GPOReport -Guid $gpo.Id -ReportType Xml)
    $links = $reportXml.GPO.LinksTo
    if ($links -eq $null -or $links.SOMPath -eq 'none') {
        $gpo | Select-Object DisplayName
    }
}

# Output the results to an Out-GridView
$excelPath = Join-Path -Path $reportsDir -ChildPath "AD_Output.xlsx"
$unlinkedGPOs | Export-Excel -Path $excelPath -WorksheetName "Unlinked GPO" -AutoSize -TableName "UnlinkGPO" -TableStyle Medium12 -Append
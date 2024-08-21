
$htmlPath = Join-Path -Path $reportsDir -ChildPath "TestimoSummary.html"
#Invoke-Testimo -Source 'ForestOptionalFeatures','DomainWellKnownFolders','ForestSubnets' -Online -ReportPath $htmlPath -AlwaysShowSteps
Invoke-Testimo -Sources DomainComputersUnsupported, DomainDuplicateObjects -SplitReports -ReportPath $htmlPath -AlwaysShowSteps
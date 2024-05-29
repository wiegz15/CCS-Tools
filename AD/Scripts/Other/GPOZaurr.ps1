
$htmlPath = Join-Path -Path $reportsDir -ChildPath "GPOList.html"

Invoke-GPOZaurr -Type GPOList -FilePath $htmlPath -HideSteps
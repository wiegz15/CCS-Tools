








$htmlPath = Join-Path -Path $reportsDir -ChildPath "TestimoSummary.html"

$Sources = @(
    'ForestRoles'
    'ForestOptionalFeatures'
    #'ForestOrphanedAdmins'
    'DomainPasswordComplexity'
    'DomainKerberosAccountAge'
    'DomainDNSScavengingForPrimaryDNSServer'
    'DomainSysVolDFSR'
    'DCRDPSecurity'
    'DCSMBShares'
    'DomainGroupPolicyMissingPermissions'
    'DCWindowsRolesAndFeatures'
    'DCNTDSParameters'
    'DCInformation'
    'ForestReplicationStatus'
)

Invoke-Testimo -Sources $Sources -ReportPath $htmlPath -AlwaysShowSteps
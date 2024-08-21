# Import the Active Directory module
Import-Module ActiveDirectory

# Load Windows Forms assembly for GUI interaction
Add-Type -AssemblyName System.Windows.Forms

# Get all server objects from Active Directory
$servers = Get-ADComputer -Filter {OperatingSystem -Like "*Server*"} -Property Name, DistinguishedName

# Prepare an array to store the results
$results = @()

foreach ($server in $servers) {
    if ($server.DistinguishedName -notmatch 'OU=.*disabled') {
        Write-Host "Checking server: $($server.Name)"
        $result = New-Object -TypeName PSObject
        $result | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $server.Name
        $ou = ($server.DistinguishedName -split ',',2)[1]
        $result | Add-Member -MemberType NoteProperty -Name "OrganizationalUnit" -Value $ou

        Write-Host "Attempting DNS check for $($server.Name)..."
        try {
            $dnsResult = Resolve-DnsName $server.Name -ErrorAction Stop
            $result | Add-Member -MemberType NoteProperty -Name "DNS" -Value "Found"
            $result | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value $dnsResult.IPAddress
            Write-Host "DNS check successful for $($server.Name), IP: $($dnsResult.IPAddress)"
        } catch {
            $result | Add-Member -MemberType NoteProperty -Name "DNS" -Value "Not Found"
            $result | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value "N/A"
            Write-Host "DNS check failed for $($server.Name)"
        }

        $connectionTest = Test-Connection $server.Name -Count 1 -Quiet
        if ($connectionTest) {
            $result | Add-Member -MemberType NoteProperty -Name "Online" -Value "Yes"
        } else {
            $result | Add-Member -MemberType NoteProperty -Name "Online" -Value "No"
        }
        $results += $result
    } else {
        Write-Host "Skipping server $($server.Name) as it is in a disabled OU."
    }
}

# Output results to GridView and allow selection of servers to delete
$selectedServers = $results | Out-GridView -Title "Server Status Report" -PassThru

# Confirm deletion of selected servers using a GUI
foreach ($server in $selectedServers) {
    $messageBoxResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete $($server.ServerName) from AD?", "Confirm Deletion", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    if ($messageBoxResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Remove-ADComputer -Identity $server.ServerName -Confirm:$false
        Write-Host "$($server.ServerName) has been deleted from Active Directory."
    } else {
        Write-Host "Deletion cancelled for $($server.ServerName)."
    }
}

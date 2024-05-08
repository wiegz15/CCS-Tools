# Load Windows Forms and drawing libraries
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Snapshot Management'
$form.Size = New-Object System.Drawing.Size(320, 320) # Adjusted form size
$form.StartPosition = 'CenterScreen'

# Snapshot Age field
$labelAge = New-Object System.Windows.Forms.Label
$labelAge.Location = New-Object System.Drawing.Point(10, 20)
$labelAge.Size = New-Object System.Drawing.Size(280, 20)
$labelAge.Text = 'Snapshot Age in Days:'
$form.Controls.Add($labelAge)

$textAge = New-Object System.Windows.Forms.TextBox
$textAge.Location = New-Object System.Drawing.Point(10, 45)
$textAge.Size = New-Object System.Drawing.Size(280, 20)
$textAge.Text = '0' # Default value
$form.Controls.Add($textAge)

# Submit Button
$submit = New-Object System.Windows.Forms.Button
$submit.Location = New-Object System.Drawing.Point(10, 75)
$submit.Size = New-Object System.Drawing.Size(280, 30)
$submit.Text = 'Submit'
$form.Controls.Add($submit)

$submit.Add_Click({
    # Find old snapshots and include the VM name
    $snapshots = Get-VM | Get-Snapshot | Where-Object { $_.Created -lt (Get-Date).AddDays(-[int]$textAge.Text) } | ForEach-Object {
        # Create a custom PSObject that includes VM Name, Snapshot Name, Created Date, and Size
        [PSCustomObject]@{
            VMName = $_.VM.Name
            SnapshotName = $_.Name
            Created = $_.Created
            SizeMB = $_.SizeMB
        }
    }

    # Display snapshots in Out-GridView, now including VM names
    $selectedSnapshots = $snapshots | Out-GridView -PassThru -Title "Select Snapshots to Delete"

    # If no snapshots were selected, close the form
    if (!$selectedSnapshots -or $selectedSnapshots.Count -eq 0) {
        $form.Close()
    } elseif ([System.Windows.Forms.MessageBox]::Show("Do you want to delete the selected snapshots?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes') {
        foreach ($snapshot in $selectedSnapshots) {
            # Retrieve the snapshot by VM and snapshot name due to how Out-GridView passes objects back
            $vm = Get-VM -Name $snapshot.VMName
            $snapshotToDelete = Get-Snapshot -VM $vm | Where-Object { $_.Name -eq $snapshot.SnapshotName }

            # Delete the snapshot
            $snapshotToDelete | Remove-Snapshot -Confirm:$false
            Write-Output "Deleted snapshot: $($snapshot.SnapshotName) for VM: $($snapshot.VMName)"
        }

        # Display summary of actions taken
        [System.Windows.Forms.MessageBox]::Show("Snapshots deletion process completed.")
    }
})

$form.ShowDialog()

Import-Module ActiveDirectory
Import-Module ImportExcel

# Get all Domain Controllers
$domainControllers = Get-ADDomainController -Filter *

$systemHealth = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    $dcName = $dc.HostName

    # Get CPU Load
    $cpuLoad = Get-WmiObject -Class Win32_Processor -ComputerName $dcName | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average

    # Get Memory Usage
    $memory = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $dcName
    $totalMemory = $memory.TotalVisibleMemorySize
    $freeMemory = $memory.FreePhysicalMemory
    $usedMemory = $totalMemory - $freeMemory
    $memoryUsagePercentage = [math]::Round(($usedMemory / $totalMemory) * 100, 2)

    # Get Disk Space
    $diskSpace = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $dcName
    $diskInfo = $diskSpace | ForEach-Object {
        [PSCustomObject]@{
            DriveLetter = $_.DeviceID
            FreeSpaceGB = [math]::Round($_.FreeSpace / 1GB, 2)
            TotalSpaceGB = [math]::Round($_.Size / 1GB, 2)
        }
    }
    
    foreach ($disk in $diskInfo) {
        $systemHealth += [PSCustomObject]@{
            DCName               = $dcName
            CPULoad              = [math]::Round($cpuLoad, 2)
            MemoryUsagePercentage= $memoryUsagePercentage
            DriveLetter          = $disk.DriveLetter
            FreeSpaceGB          = $disk.FreeSpaceGB
            TotalSpaceGB         = $disk.TotalSpaceGB
        }
    }
}

# Export to Excel
$excelPath = "C:\Path\To\Your\Directory\AD_Output.xlsx"
$systemHealth | Export-Excel -Path $excelPath -WorksheetName "System Health" -AutoSize -TableName "SystemHealth" -TableStyle Medium13 -Append

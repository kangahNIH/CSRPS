# Disks Usage ALL VMs with Connection Status report 
# Make sure all VMs on ALL.txt are turn-On
# If Connectivity Errors, Test Test-NetConnection -ComputerName $server -Port 135

# Load server list
$servers = Get-Content "O:\All.txt"
$results = @()

foreach ($server in $servers) {
    try {
        # Step 1: Ping test
        $ping = Test-Connection -ComputerName $server -Count 1 -Quiet
        if (-not $ping) {
            throw "Ping failed - server may be offline"
        }

        # Step 2: WMI connectivity test
        $wmiCheck = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction Stop

        # Step 3: Query logical disks (DriveType=3 means fixed disks)
        $disks = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $server -Filter "DriveType=3" -ErrorAction Stop

        foreach ($disk in $disks) {
            $totalGB = [math]::Round($disk.Size / 1GB, 2)
            $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
            $usedPercent = if ($disk.Size -ne 0) {
                [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 2)
            } else {
                0
            }

            $results += [PSCustomObject]@{
                Server      = $server
                Drive       = $disk.DeviceID
                TotalGB     = $totalGB
                FreeGB      = $freeGB
                UsedPercent = "$usedPercent%"
                Status      = "Online"
            }
        }
    } catch {
        # Server unreachable or WMI failed
        $results += [PSCustomObject]@{
            Server      = $server
            Drive       = "N/A"
            TotalGB     = "N/A"
            FreeGB      = "N/A"
            UsedPercent = "N/A"
            Status      = "Offline or Unreachable"
        }
    }
}

# Export the results to CSV
$results | Export-Csv "O:\Report\DiskUsage-8-8-2025.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Disk usage report completed. Results saved to DiskUsage-8-8-2025.csv"

# Memory Usage ALL VMs with Connection Status report
# Make sure all VMs on ALL.txt are turned on
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

        # Step 3: Query memory usage
        $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction Stop
        $totalMemoryGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $freeMemoryGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
        $usedMemoryGB = $totalMemoryGB - $freeMemoryGB
        $usedPercent = if ($totalMemoryGB -ne 0) {
            [math]::Round(($usedMemoryGB / $totalMemoryGB) * 100, 2)
        } else {
            0
        }

        $results += [PSCustomObject]@{
            Server        = $server
            TotalMemoryGB = $totalMemoryGB
            FreeMemoryGB  = $freeMemoryGB
            UsedMemoryGB  = $usedMemoryGB
            UsedPercent   = "$usedPercent%"
            Status        = "Online"
        }
    } catch {
        # Server unreachable or WMI failed
        $results += [PSCustomObject]@{
            Server        = $server
            TotalMemoryGB = "N/A"
            FreeMemoryGB  = "N/A"
            UsedMemoryGB  = "N/A"
            UsedPercent   = "N/A"
            Status        = "Offline or Unreachable"
        }
    }
}

# Export the results to CSV
$results | Export-Csv "O:\Report\MemoryUsage-8-8-2025.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Memory usage report completed. Results saved to MemoryUsage-8-8-2025.csv"

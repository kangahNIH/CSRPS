# Memory Usage - AVG ALL VMs with Connection Status report
# Includes average memory usage over 3 samples

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

        # Step 3: Sample memory usage 3 times
        $samples = @()
        for ($i = 0; $i -lt 3; $i++) {
            $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction Stop
            $total = $os.TotalVisibleMemorySize / 1MB
            $free  = $os.FreePhysicalMemory / 1MB
            $used  = $total - $free
            $samples += [PSCustomObject]@{ Total = $total; Used = $used }
            Start-Sleep -Seconds 2
        }

        $avgTotal = [math]::Round(($samples | Measure-Object -Property Total -Average).Average, 2)
        $avgUsed  = [math]::Round(($samples | Measure-Object -Property Used -Average).Average, 2)
        $avgFree  = [math]::Round($avgTotal - $avgUsed, 2)
        $avgPercent = if ($avgTotal -ne 0) {
            [math]::Round(($avgUsed / $avgTotal) * 100, 2)
        } else {
            0
        }

        $results += [PSCustomObject]@{
            Server         = $server
            AvgTotalGB     = $avgTotal
            AvgUsedGB      = $avgUsed
            AvgFreeGB      = $avgFree
            AvgUsedPercent = "$avgPercent%"
            Status         = "Online"
        }
    } catch {
        # Server unreachable or WMI failed
        $results += [PSCustomObject]@{
            Server         = $server
            AvgTotalGB     = "N/A"
            AvgUsedGB      = "N/A"
            AvgFreeGB      = "N/A"
            AvgUsedPercent = "N/A"
            Status         = "Offline or Unreachable"
        }
    }
}

# Export the results to CSV
$results | Export-Csv "O:\Report\MemoryUsageAvgReport.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Average memory usage report completed. Results saved to MemoryUsageAvgReport.csv"

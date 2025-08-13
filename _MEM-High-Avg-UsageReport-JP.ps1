# Memory Usage Report for All VMs
# Includes average and peak memory usage (GB and %)

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
        $osCheck = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction Stop

        ### Memory Section ###
        $memSamples = @()
        for ($i = 0; $i -lt 5; $i++) {
            $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server -ErrorAction Stop
            $totalMem = $os.TotalVisibleMemorySize / 1MB
            $freeMem  = $os.FreePhysicalMemory / 1MB
            $usedMem  = $totalMem - $freeMem
            $memSamples += [PSCustomObject]@{ Total = $totalMem; Used = $usedMem }
            Start-Sleep -Seconds 2
        }

        # Average usage
        $avgUsedGB = [math]::Round(($memSamples | Measure-Object -Property Used -Average).Average, 2)
        $avgTotalGB = [math]::Round(($memSamples | Measure-Object -Property Total -Average).Average, 2)
        $avgPercent = if ($avgTotalGB -ne 0) {
            [math]::Round(($avgUsedGB / $avgTotalGB) * 100, 2)
        } else { 0 }

        # Peak usage
        $peakUsedGB = [math]::Round(($memSamples | Measure-Object -Property Used -Maximum).Maximum, 2)
        $peakPercent = if ($avgTotalGB -ne 0) {
            [math]::Round(($peakUsedGB / $avgTotalGB) * 100, 2)
        } else { 0 }

        ### Final Result ###
        $results += [PSCustomObject]@{
            Server            = $server
            AvgMemUsedGB      = "$avgUsedGB GB"
            AvgMemUsagePct    = "$avgPercent%"
            PeakMemUsedGB     = "$peakUsedGB GB"
            PeakMemUsagePct   = "$peakPercent%"
            Status            = "Online"
        }
    } catch {
        # Server unreachable or WMI failed
        $results += [PSCustomObject]@{
            Server            = $server
            AvgMemUsedGB      = "N/A"
            AvgMemUsagePct    = "N/A"
            PeakMemUsedGB     = "N/A"
            PeakMemUsagePct   = "N/A"
            Status            = "Offline or Unreachable"
        }
    }
}

# Export the results to CSV
$results | Export-Csv "O:\Report\Memory_Usage-High-Avg-Report.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Memory usage report completed. Results saved to Memory_UsageReport.csv"

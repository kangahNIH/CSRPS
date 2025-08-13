# CPU Usage Report High-Avg for All VMs
# Includes average and high CPU usage %, core count, and connection status

$servers = Get-Content "C:\STPS\All.txt"
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

        ### CPU Section ###
        # Get CPU core count
        $cpuInfo = Get-WmiObject -Class Win32_Processor -ComputerName $server -ErrorAction Stop
        $coreCount = ($cpuInfo | Measure-Object -Property NumberOfCores -Sum).Sum

        # Sample CPU usage 3 times
        $cpuSamples = Get-Counter -ComputerName $server '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 3
        $avgCpu = [math]::Round(($cpuSamples.CounterSamples | Measure-Object -Property CookedValue -Average).Average, 2)

        ### Final Result ###
        $results += [PSCustomObject]@{
            Server         = $server
            CPUCoreCount   = $coreCount
            AvgCPUUsagePct = "$avgCpu%"
            Status         = "Online"
        }
    } catch {
        # Server unreachable or WMI failed
        $results += [PSCustomObject]@{
            Server         = $server
            CPUCoreCount   = "N/A"
            AvgCPUUsagePct = "N/A"
            Status         = "Offline or Unreachable"
        }
    }
}

# Export the results to CSV
$results | Export-Csv "C:\STPS\Report\CPU_Usage_High-Avg-Report.csv" -NoTypeInformation -Encoding UTF8
Write-Host "CPU usage report completed. Results saved to CPU_UsageReport.csv"

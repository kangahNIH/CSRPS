# CPU Usage ALL VMs with Connection Status report
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

        # Step 3: Get CPU core count
        $cpuInfo = Get-WmiObject -Class Win32_Processor -ComputerName $server -ErrorAction Stop
        $coreCount = ($cpuInfo | Measure-Object -Property NumberOfCores -Sum).Sum

        # Step 4: Get CPU usage via performance counter
        $cpuUsage = Get-Counter -ComputerName $server '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 3
        $avgCpu = [math]::Round(($cpuUsage.CounterSamples | Measure-Object -Property CookedValue -Average).Average, 2)

        $results += [PSCustomObject]@{
            Server     = $server
            CPUUsage   = "$avgCpu%"
            CoreCount  = $coreCount
            Status     = "Online"
        }
    } catch {
        # Server unreachable or WMI failed
        $results += [PSCustomObject]@{
            Server     = $server
            CPUUsage   = "N/A"
            CoreCount  = "N/A"
            Status     = "Offline or Unreachable"
        }
    }
}

# Export the results to CSV
$results | Export-Csv "O:\Report\CPUUsage-8-8-2025.csv" -NoTypeInformation -Encoding UTF8
Write-Host "CPU usage report completed. Results saved to CPUUsage-8-8-2025.csv"

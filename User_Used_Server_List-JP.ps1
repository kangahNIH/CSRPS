$StartDate = (Get-Date).AddMonths(-3)
$Computers = Get-Content "C:\STPS\GREEN.txt"
$OutputFile = "O:\reports\UseruseServers_3month_6-3-2025.csv"

$Results = @()

foreach ($Computer in $Computers) {
    try {
        $Events = Get-EventLog -LogName Security -ComputerName $Computer | Where-Object {
            $_.EventID -eq 4624 -and $_.TimeGenerated -ge $StartDate
        }

        foreach ($Event in $Events) {
            $UserInfo = [PSCustomObject]@{
                ComputerName  = $Computer
                TimeGenerated = $Event.TimeGenerated
                UserDetails   = $Event.Message
            }
            $Results += $UserInfo
        }
    } catch {
        Write-Warning "Failed to retrieve logs from $Computer"
    }
}

$Results | Export-Csv -Path $OutputFile -NoTypeInformation

Write-Host "User login data from the past 3 months saved to $OutputFile"

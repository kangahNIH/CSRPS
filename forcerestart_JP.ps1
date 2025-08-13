# forcerestart_JP.ps1

# --- Configuration ---
# Path to your list of computer names (one per line)
$ComputerListFile = "o:\ps\test2.txt"

# --- Script Starts Here ---

Write-Host "Starting the computer restart process..."

# Check if the file with computer names exists
if (-not (Test-Path $ComputerListFile)) {
    Write-Error "Error: The file '$ComputerListFile' was not found. Please check the path."
    exit # Stop the script if the file isn't there
}

# Read computer names from the file, skipping empty lines
$ComputersToRestart = Get-Content $ComputerListFile | Where-Object { $_.Trim() -ne "" }

if ($ComputersToRestart.Count -eq 0) {
    Write-Warning "No computer names found in the file. Nothing to do!"
    exit # Stop if the file is empty
}

Write-Host "Found $($ComputersToRestart.Count) computer(s) to restart."

# Go through each computer in the list
foreach ($Computer in $ComputersToRestart) {
    Write-Host "`n--- Attempting to restart $($Computer) ---"

    try {
        # This is the command that runs on the remote computer:
        # 1. Stop all running applications forcefully
        # 2. Restart the computer forcefully
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            Write-Host "Stopping apps and restarting..."
            Stop-Process -Name * -Force -ErrorAction SilentlyContinue
            Restart-Computer -Force -Confirm:$false
        } `
        -ErrorAction Stop ` # If Invoke-Command itself fails, show the error
        -AsJob:$false # Run one by one

        Write-Host "Successfully sent restart command to $($Computer)." -ForegroundColor Green
    }
    catch {
        # If something went wrong, tell us what happened
        Write-Error "Failed to restart $($Computer). Error: $($_.Exception.Message)"
        Write-Host "Possible issues: Computer offline, network problem, or permissions." -ForegroundColor Red
    }

    # Wait a bit before trying the next computer
    Start-Sleep -Seconds 3
}

Write-Host "`nAll specified computers have been processed."
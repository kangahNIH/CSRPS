# Define variables
$psexecPath = "C:\Temp\psexec.exe"
$serversFile = "O:\TEST4.txt"
$installerSourcePath = "C:\Temp\CohesityAgent\Cohesity_Agent_7.1.2_u3_20241231_Win_x64_Installer.exe"
$installerDestinationPath = "C:\Temp\CohesityAgentInstaller.exe"
$reportFile = "O:\Report\CohesityAgentInstalled_8-1-2025.csv"

# Create a custom object to store results
$results = @()

# --- Initial File Checks ---

# Check if the PsExec executable exists
if (-not (Test-Path $psexecPath)) {
    Write-Host "Error: PsExec.exe was not found at '$psexecPath'."
    Write-Host "Please download the PSTools suite and place psexec.exe at this location."
    exit
}

# Check if the servers file exists
if (-not (Test-Path $serversFile)) {
    Write-Host "Error: The server list file '$serversFile' was not found."
    exit
}

# Check if the installer file exists
if (-not (Test-Path $installerSourcePath)) {
    Write-Host "Error: The installer file '$installerSourcePath' was not found."
    exit
}

# Read the list of servers from the file and filter out comment lines
$servers = Get-Content $serversFile | Where-Object { -not $_.TrimStart().StartsWith('#') }

# --- Loop through each server ---

foreach ($server in $servers) {
    Write-Host "Processing server: $server"

    # Reset status and error message for each loop
    $status = "Unknown"
    $errorMessage = ""

    try {
        # Check if the server is online
        if (-not (Test-Connection -ComputerName $server -Count 1 -Quiet)) {
            $status = "Offline"
            $errorMessage = "Server is not reachable via ping."
            Write-Host "Server '$server' is offline. Skipping..."
            # Continue to the next server in the list
            continue
        }

        # Copy the installer to the remote server using the admin share
        Write-Host "Copying installer to $server..."
        # The 'c$' share is a standard administrative share, accessible with admin rights.
        Copy-Item -Path $installerSourcePath -Destination "\\$server\c$\Temp\CohesityAgentInstaller.exe" -Force -ErrorAction Stop

        # Run the installer remotely using PsExec
        Write-Host "Executing installer on $server using PsExec..."

        # The command to execute:
        # psexec.exe \\server -s -accepteula cmd.exe /c "C:\Temp\CohesityAgentInstaller.exe /s"
        # -s: runs the process as the System account
        # -accepteula: accepts the EULA to prevent prompts
        # cmd.exe /c "..." : this is a reliable way to run the command and get its exit code.
        $arguments = "\\$server -s -accepteula cmd.exe /c ""C:\Temp\CohesityAgentInstaller.exe /s"""

        # Start the PsExec process and wait for it to complete.
        $process = Start-Process -FilePath $psexecPath -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        $exitCode = $process.ExitCode

        # Analyze the exit code. PsExec will return the exit code of the remote process.
        if ($exitCode -eq 0) {
            $status = "Success"
            Write-Host "Installation on '$server' was successful."
        } else {
            $status = "Failed"
            $errorMessage = "Installation failed with exit code: $exitCode"
            Write-Host "$errorMessage"
        }

    }
    catch {
        $status = "Failed"
        $errorMessage = $_.Exception.Message
        Write-Host "An error occurred while processing server '$server': $errorMessage"
    }
    finally {
        # Add the result to the array regardless of success or failure
        $results += [PSCustomObject]@{
            ServerName = $server
            InstallationStatus = $status
            ErrorMessage = $errorMessage
        }
    }
}

# Export the results to a CSV file
$results | Export-Csv -Path $reportFile -NoTypeInformation

Write-Host "Script completed. Results exported to '$reportFile'."
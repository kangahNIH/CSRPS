# Define variables
$psexecPath = "C:\STPS\psexec.exe" # Path to psexec.exe
$serversFile = "O:\TEST4.txt" # Path to your list of servers
$installerSourcePath = "C:\STPS\Cohesity_Agent_7.1.2_u3_20241231_Win_x64_Installer.exe" # Local path to the installer
$installerDestinationPath = "C:\Temp\CohesityAgentInstaller.exe" # Consistent destination path on the remote server
$reportFile = "O:\Report\CohesityAgentInstalled_8-1-2025.csv"

# Create a custom object to store results
$results = @()

# --- Initial File Checks ---
if (-not (Test-Path $psexecPath)) {
    Write-Host "Error: PsExec.exe was not found at '$psexecPath'."
    exit
}
if (-not (Test-Path $serversFile)) {
    Write-Host "Error: The server list file '$serversFile' was not found."
    exit
}
if (-not (Test-Path $installerSourcePath)) {
    Write-Host "Error: The installer file '$installerSourcePath' was not found."
    exit
}

# Read the list of servers, filtering out comments and blank lines
$servers = Get-Content $serversFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.TrimStart().StartsWith('#') }

# --- Loop through each server ---
foreach ($server in $servers) {
    Write-Host "Processing server: $server"

    $status = "Unknown"
    $errorMessage = ""

    try {
        if (-not (Test-Connection -ComputerName $server -Count 1 -Quiet)) {
            $status = "Offline"
            $errorMessage = "Server is not reachable via ping."
            Write-Host "Server '$server' is offline. Skipping..."
            continue
        }

        # First, attempt to remove any old installer file to prevent file lock errors
        Write-Host "Attempting to remove old installer file on $server..."
        Remove-Item -Path "\\$server\c$\Temp\CohesityAgentInstaller.exe" -ErrorAction SilentlyContinue

        # Copy the installer to the remote server using the admin share
        Write-Host "Copying installer to $server..."
        Copy-Item -Path $installerSourcePath -Destination "\\$server\c$\Temp\CohesityAgentInstaller.exe" -Force -ErrorAction Stop

        # Run the installer remotely using PsExec
        Write-Host "Executing installer on $server using PsExec..."
        # --- NEW SILENT INSTALL COMMAND ---
        # The /s /v"/qn" flags are a more robust way to force a completely silent installation.
        $arguments = "\\$server -s -accepteula cmd.exe /c ""C:\Temp\CohesityAgentInstaller.exe /s /v""/qn"""""
        $process = Start-Process -FilePath $psexecPath -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        $exitCode = $process.ExitCode

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
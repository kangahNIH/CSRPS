# Define variables
$serversFile = "O:\TEST.txt"
$installerSourcePath = "C:\Temp\CohesityAgent\Cohesity_Agent_7.1.2_u3_20241231_Win_x64_Installer.exe"
$installerDestinationPath = "C:\Temp\CohesityAgentInstaller.exe"
$reportFile = "O:\Report\CohesityAgentInstalled_8-1-2025.csv"

# Create a custom object to store results
$results = @()

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

# Loop through each server
foreach ($server in $servers) {
    Write-Host "Processing server: $server"

    # Reset status and error message for each loop
    $status = "Unknown"
    $errorMessage = ""
    $session = $null

    try {
        # Check if the server is online
        if (-not (Test-Connection -ComputerName $server -Count 1 -Quiet)) {
            $status = "Offline"
            $errorMessage = "Server is not reachable."
            Write-Host "Server '$server' is offline. Skipping..."
            continue
        }

        # Create a remote session to the server
        Write-Host "Attempting to create PSSession to $server..."
        # Add -ErrorAction Stop to make sure the PSSession failure is caught immediately
        $session = New-PSSession -ComputerName $server -ErrorAction Stop

        # Copy the installer to the remote server
        Write-Host "Copying installer to $server..."
        # Add -ErrorAction Stop here as well
        Copy-Item -Path $installerSourcePath -Destination "\\$server\c$\Temp\CohesityAgentInstaller.exe" -Force -ErrorAction Stop

        # Run the installer with silent switches and capture the exit code
        Write-Host "Executing installer on $server..."
        $scriptBlock = {
            $installFile = "C:\Temp\CohesityAgentInstaller.exe"
            $process = Start-Process -FilePath $installFile -ArgumentList '/s' -Wait -PassThru
            return $process.ExitCode
        }
        $exitCode = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ErrorAction Stop

        # Analyze the exit code
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
        # Add the result to the array
        $results += [PSCustomObject]@{
            ServerName = $server
            InstallationStatus = $status
            ErrorMessage = $errorMessage
        }

        # Close the remote session if it was successfully created
        if ($session) {
            Write-Host "Closing PSSession to $server..."
            Remove-PSSession -Session $session
        }
    }
}

# Export the results to a CSV file
$results | Export-Csv -Path $reportFile -NoTypeInformation

Write-Host "Script completed. Results exported to '$reportFile'."
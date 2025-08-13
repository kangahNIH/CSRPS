Import-Module ActiveDirectory

# Prompt for privileged AD credentials
$credential = Get-Credential -UserName "aakangah@nih.gov" -Message "Enter password for privileged AD account"

# Define AD Domain
$domain = "nih.gov"

# CSV paths
$inputCsv = "C:\STPS\emails_CSR_Incoming_SRO.csv"
$outputCsv = "C:\STPS\AD_UserDetails_$(Get-Date -Format 'MMdd_yyyy_HHmm').csv"

# Import CSV
$emails = Import-Csv -Path $inputCsv

# Initialize results
$results = @()

foreach ($emailEntry in $emails) {
    # Safely handle potential empty or null entries
    if ($emailEntry.PSObject.Properties.Name -contains "Email" -and `
        -not [string]::IsNullOrWhiteSpace($emailEntry.Email)) {
        
        $email = $emailEntry.Email.Trim()

        try {
            # Query AD user details
            $user = Get-ADUser -Filter "Mail -eq '$email'" `
                -Properties SamAccountName, Mail, Department `
                -Server $domain `
                -Credential $credential

            if ($user) {
                $results += [PSCustomObject]@{
                    SamAccountName = $user.SamAccountName
                    EmailAddress   = $user.Mail
                    Department     = $user.Department
                }
            }
            else {
                Write-Warning "No AD user found for email: $email"
            }
        }
        catch {
            Write-Warning "AD lookup error for email '$email': $_"
        }
    }
    else {
        Write-Warning "Invalid or empty email entry detected and skipped."
    }
}

# Export results
$results | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Export completed successfully: $outputCsv" -ForegroundColor Green

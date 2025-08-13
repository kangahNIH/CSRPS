# Check and install Active Directory Module if missing
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Installing Active Directory Module..."
    Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature -IncludeManagementTools
}

Import-Module ActiveDirectory

# Prompt securely for AD credentials
$Credential = Get-Credential -Message "Enter your AD Credentials"

# Define groups and DLs with distinguished names
$groups = @(
    @{Name = "CSR_ALL"; DN = "CN=CSR All,OU=Distribution Lists,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Staff"; DN = "CN=CSR Staff,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_SRO"; DN = "CN=CSR SRO,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Review_Group"; DN = "CN=CSR Review,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Review_DL"; DN = "CN=CSR Review,OU=Distribution Lists,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"}
)

# Directory for CSV outputs updated to C:\STPS
$outputDir = "C:\STPS"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Array to collect non-CSR department users
$nonCSRUsers = @()

# Process each group and export to individual CSV files
foreach ($group in $groups) {
    Write-Host "Processing $($group.Name)..."

    # Get group members
    $members = Get-ADGroupMember -Identity $group.DN -Recursive -Credential $Credential |
        Where-Object { $_.ObjectClass -eq 'user' } |
        Get-ADUser -Properties samAccountName, Surname, GivenName, Department, EmailAddress -Credential $Credential

    # Separate CSR and Non-CSR users
    $CSRUsers = @()
    foreach ($user in $members) {
        $userData = [PSCustomObject]@{
            SamAccountName = $user.samAccountName
            LastName       = $user.Surname
            FirstName      = $user.GivenName
            EmailAddress   = $user.EmailAddress
            Department     = $user.Department
        }

        if ($user.Department -eq "CSR") {
            $CSRUsers += $userData
        }
        else {
            # Add to Non-CSR users list
            $userData | Add-Member -NotePropertyName Group -NotePropertyValue $group.Name
            $nonCSRUsers += $userData
        }
    }

    # Export CSR users to CSV
    $csvPath = Join-Path $outputDir ("$($group.Name)_CSR.csv")
    $CSRUsers | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported CSR users to $csvPath"
}

# Export Non-CSR users (from all groups) into one separate CSV
if ($nonCSRUsers.Count -gt 0) {
    $nonCSRPath = Join-Path $outputDir "Non_CSR_Users.csv"
    $nonCSRUsers | Export-Csv -Path $nonCSRPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nNon-CSR users exported to: $nonCSRPath" -ForegroundColor Yellow
}
else {
    Write-Host "`nNo Non-CSR users found." -ForegroundColor Green
}

Write-Host "`nCSV generation completed successfully."

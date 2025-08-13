# Ensure AD Module is installed
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Installing Active Directory Module..."
    Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature -IncludeManagementTools
}
Import-Module ActiveDirectory

# Prompt securely for AD credentials
$Credential = Get-Credential -Message "Enter your AD Credentials"

# Define groups and DLs
$groups = @(
    @{Name = "CSR_ALL"; DN = "CN=CSR All,OU=Distribution Lists,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Staff"; DN = "CN=CSR Staff,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_SRO"; DN = "CN=CSR SRO,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Review_Group"; DN = "CN=CSR Review,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Review_DL"; DN = "CN=CSR Review,OU=Distribution Lists,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"}
)

# Set the specific Domain Controller
$DomainController = "nih.gov"

# CSV Output Directory
$outputDir = "C:\STPS"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Get current date/time for file naming
$timestamp = Get-Date -Format "MMdd_hhmmt\\M"  # e.g., 0515_1030AM

# Array for non-CSR users
$nonCSRUsers = @()

foreach ($group in $groups) {
    Write-Host "Processing $($group.Name) from $DomainController..."

    # Fetch members explicitly from specific Domain Controller
    $members = Get-ADGroupMember -Identity $group.DN -Recursive -Credential $Credential -Server $DomainController |
        Where-Object { $_.ObjectClass -eq 'user' } |
        Get-ADUser -Properties samAccountName, Surname, GivenName, Department, EmailAddress -Credential $Credential -Server $DomainController

    # Export full current member list to CSV
    $groupMembers = $members | Select-Object `
        @{Name='SamAccountName';Expression={$_.samAccountName}}, `
        @{Name='LastName';Expression={$_.Surname}}, `
        @{Name='FirstName';Expression={$_.GivenName}}, `
        @{Name='EmailAddress';Expression={$_.EmailAddress}}, `
        @{Name='Department';Expression={$_.Department}}

    $groupCsvPath = Join-Path $outputDir ("$($group.Name)_AllMembers_$timestamp.csv")
    $groupMembers | Export-Csv -Path $groupCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "All members exported to $groupCsvPath"

    # Identify and collect Non-CSR users
    foreach ($user in $members) {
        if ($user.Department -ne "CSR") {
            $nonCSRUsers += [PSCustomObject]@{
                Group          = $group.Name
                SamAccountName = $user.samAccountName
                LastName       = $user.Surname
                FirstName      = $user.GivenName
                EmailAddress   = $user.EmailAddress
                Department     = $user.Department
            }
        }
    }
}

# Export combined Non-CSR users to separate CSV
if ($nonCSRUsers.Count -gt 0) {
    $nonCSRPath = Join-Path $outputDir "Non_CSR_Users_$timestamp.csv"
    $nonCSRUsers | Export-Csv -Path $nonCSRPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nNon-CSR users exported to: $nonCSRPath" -ForegroundColor Yellow
}
else {
    Write-Host "`nNo Non-CSR users found across all groups." -ForegroundColor Green
}

Write-Host "`nCSV generation completed successfully from Domain Controller: $DomainController"

# Ensure AD Module is installed
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Installing Active Directory Module..."
    Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature -IncludeManagementTools
}
Import-Module ActiveDirectory

# Prompt securely for AD credentials
$Credential = Get-Credential -Message "Enter your AD Credentials"

# Define group names (DNs will be pulled automatically)
$groupNames = @(
#    "CSR External Users",
    "CSR Incoming SRO"
#    "CSR Incoming Support Staff"
)

# Domain controller explicitly set
$DomainController = "nih.gov"

# Output Directory
$outputDir = "C:\STPS"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Timestamp for filenames
$timestamp = Get-Date -Format "MMdd_hhmmtm"

# Collect Non-CSR users across all groups
$nonCSRUsers = @()

foreach ($groupName in $groupNames) {
    Write-Host "Processing group: $groupName from $DomainController..."

    # Retrieve the group Distinguished Name dynamically
    $group = Get-ADGroup -Filter "Name -eq '$groupName'" -Credential $Credential -Server $DomainController
    if (-not $group) {
        Write-Warning "Group '$groupName' not found. Skipping."
        continue
    }

    # Get group members
    $members = Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -Credential $Credential -Server $DomainController |
        Where-Object { $_.ObjectClass -eq 'user' } |
        Get-ADUser -Properties samAccountName, Surname, GivenName, Department, EmailAddress -Credential $Credential -Server $DomainController

    # Prepare data sorted by Last Name
    $sortedMembers = $members | Sort-Object Surname, GivenName | Select-Object `
        @{Name='SamAccountName';Expression={$_.samAccountName}}, `
        @{Name='LastName';Expression={$_.Surname}}, `
        @{Name='FirstName';Expression={$_.GivenName}}, `
        @{Name='EmailAddress';Expression={$_.EmailAddress}}, `
        @{Name='Department';Expression={$_.Department}}

    # CSV output path
    $groupFileName = ($groupName -replace ' ', '_') + "_AllMembers_$timestamp.csv"
    $groupCsvPath = Join-Path $outputDir $groupFileName

    # Export sorted members to CSV
    $sortedMembers | Export-Csv -Path $groupCsvPath -NoTypeInformation -Encoding UTF8

    # Add total count as last line
    Add-Content -Path $groupCsvPath -Value ""
    Add-Content -Path $groupCsvPath -Value "Total Members:,$($sortedMembers.Count)"

    Write-Host "Exported sorted members to: $groupCsvPath (Total: $($sortedMembers.Count))" -ForegroundColor Green

    # Collect Non-CSR users
    foreach ($user in $members) {
        if ($user.Department -ne "CSR") {
            $nonCSRUsers += [PSCustomObject]@{
                Group          = $groupName
                SamAccountName = $user.samAccountName
                LastName       = $user.Surname
                FirstName      = $user.GivenName
                EmailAddress   = $user.EmailAddress
                Department     = $user.Department
            }
        }
    }
}

# Export combined Non-CSR users to separate CSV, sorted by LastName
if ($nonCSRUsers.Count -gt 0) {
    $nonCSRSorted = $nonCSRUsers | Sort-Object LastName, FirstName
    $nonCSRPath = Join-Path $outputDir "Non_CSR_Users_$timestamp.csv"
    $nonCSRSorted | Export-Csv -Path $nonCSRPath -NoTypeInformation -Encoding UTF8

    # Add total count as last line
    Add-Content -Path $nonCSRPath -Value ""
    Add-Content -Path $nonCSRPath -Value "Total Non-CSR Members:,$($nonCSRSorted.Count)"

    Write-Host "`nNon-CSR users exported to: $nonCSRPath (Total: $($nonCSRSorted.Count))" -ForegroundColor Yellow
}
else {
    Write-Host "`nNo Non-CSR users found across all groups." -ForegroundColor Green
}

Write-Host "`nAll CSV files generated successfully from Domain Controller: $DomainController" -ForegroundColor Cyan

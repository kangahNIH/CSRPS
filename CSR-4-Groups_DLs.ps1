# Check and install Active Directory Module if missing
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Installing Active Directory Module..."
    Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature -IncludeManagementTools
}

Import-Module ActiveDirectory

# Prompt for credentials
$Credential = Get-Credential -Message "Enter your AD Credentials"

# Define groups and DLs with distinguished names
$groups = @(
    @{Name = "CSR_ALL"; DN = "CN=CSR All,OU=Distribution Lists,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Staff"; DN = "CN=CSR Staff,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_SRO"; DN = "CN=CSR SRO,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Review_Group"; DN = "CN=CSR Review,OU=Groups,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"},
    @{Name = "CSR_Review_DL"; DN = "CN=CSR Review,OU=Distribution Lists,OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"}
)

# Directory for CSV outputs
$outputDir = "C:\STPS\AD_GroupMembers_CSV"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Process each group and export to individual CSV files
foreach ($group in $groups) {
    Write-Host "Processing $($group.Name)..."

    # Get members of the group recursively and fetch their properties
    $members = Get-ADGroupMember -Identity $group.DN -Recursive -Credential $Credential |
        Where-Object { $_.ObjectClass -eq 'user' } |
        Get-ADUser -Properties samAccountName, Surname, GivenName, Department -Credential $Credential

    # Prepare data for export
    $exportData = $members | Select-Object `
        @{Name='SamAccountName';Expression={$_.samAccountName}}, `
        @{Name='LastName';Expression={$_.Surname}}, `
        @{Name='FirstName';Expression={$_.GivenName}}, `
        @{Name='Department';Expression={$_.Department}}

    # Export data to CSV
    $csvPath = Join-Path $outputDir ("$($group.Name).csv")
    $exportData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "Exported to $csvPath"
}

Write-Host "`nAll CSV files have been successfully created in:"
Write-Host "`n$outputDir" -ForegroundColor Green

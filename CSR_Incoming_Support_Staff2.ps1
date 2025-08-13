# Ensure AD module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Install-WindowsFeature RSAT-AD-PowerShell
}
Import-Module ActiveDirectory

# Prompt for credentials
$username = "aakangah@nih.gov"
$securePassword = Read-Host -Prompt "Enter password for $username" -AsSecureString
$credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

# Input CSV and dynamic output path
$inputPath = "C:\STPS\CSRIncomingSupportStaff0519.csv"
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputPath = "C:\STPS\Matched_SamAccounts_0519_$timestamp.csv"

# Read input CSV
$inputData = Import-Csv -Path $inputPath
$results = @()

foreach ($entry in $inputData) {
    # Extract names from "LastName, FirstName" format
    if ($entry.Name -match "^\s*(.+?),\s*(.+?)\s*$") {
        $lastName = $matches[1]
        $firstName = $matches[2]
        $department = $entry.IC

        # Search AD with provided credential
        $user = Get-ADUser -Filter {
            Surname -eq $lastName -and GivenName -eq $firstName -and Department -eq $department
        } -Properties Department -Credential $credential

        if ($user) {
            $results += [PSCustomObject]@{
                samaccountname = $user.SamAccountName
                LastName       = $lastName
                FirstName      = $firstName
                Department     = $user.Department
            }
        } else {
            Write-Warning "No AD match for: $firstName $lastName in department '$department'"
        }
    } else {
        Write-Warning "Invalid name format: $($entry.Name)"
    }
}

# Export to CSV
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "✅ Output written to: $outputPath"

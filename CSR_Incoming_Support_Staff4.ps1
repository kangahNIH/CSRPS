# Ensure AD module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Install-WindowsFeature RSAT-AD-PowerShell
}
Import-Module ActiveDirectory

# Prompt for AD credentials
$username = "aakangah@nih.gov"
$securePassword = Read-Host -Prompt "Enter password for $username" -AsSecureString
$credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

# Paths
$inputPath = "C:\STPS\CSRIncomingSupportStaff0519.csv"
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputCsvPath = "C:\STPS\Matched_SamAccounts_0519_$timestamp.csv"
$outputSamListPath = "C:\STPS\SamAccountList_0519_$timestamp.txt"

# Read input
$inputData = Import-Csv -Path $inputPath
$results = @()
$samAccountNames = @()

foreach ($entry in $inputData) {
    $match = [regex]::Match($entry.Name, "^\s*(.+?),\s*(.+?)\s*$")
    if ($match.Success) {
        $lastName = $match.Groups[1].Value
        $firstName = $match.Groups[2].Value
        $expectedDept = $entry.IC.Trim()

        # Query AD for name match
        $user = Get-ADUser -Filter {
            Surname -eq $lastName -and GivenName -eq $firstName
        } -Properties Department -Credential $credential

        if ($user) {
            $actualDept = ($user.Department).Trim()

            # Always record found user's info in output CSV
            $results += [PSCustomObject]@{
                samaccountname = $user.SamAccountName
                LastName       = $lastName
                FirstName      = $firstName
                Input_IC       = $expectedDept
                Department     = $actualDept
            }

            # Only include in samAccountNames list if dept matches or contains IC
            if ($actualDept -like "*$expectedDept*") {
                $samAccountNames += $user.SamAccountName
            } else {
                Write-Warning "Department mismatch for $firstName $lastName : Input '$expectedDept' not in AD '$actualDept'"
            }
        } else {
            # Not found in AD
            $results += [PSCustomObject]@{
                samaccountname = ""
                LastName       = $lastName
                FirstName      = $firstName
                Input_IC       = $expectedDept
                Department     = ""
            }
            Write-Warning "User not found in AD: $firstName $lastName"
        }
    } else {
        Write-Warning "Invalid name format: $($entry.Name)"
    }
}

# Export full CSV with all users (matched or not)
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation -Encoding UTF8
Write-Host "✅ Detailed CSV output saved to: $outputCsvPath"

# Export samaccountname list (only matching departments)
$joinedList = ($samAccountNames -join ";")
Set-Content -Path $outputSamListPath -Value $joinedList
Write-Host "✅ Semicolon-separated list saved to: $outputSamListPath"

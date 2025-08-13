# Ensure AD module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Install-WindowsFeature RSAT-AD-PowerShell
}
Import-Module ActiveDirectory

# Input and output paths
$inputPath = "C:\STPS\CSRIncomingSupportStaff0519.csv"
$outputPath = "C:\STPS\Matched_SamAccounts_0519.csv"

# Read the CSV
$inputData = Import-Csv -Path $inputPath

# Prepare output list
$results = @()

foreach ($entry in $inputData) {
    # Extract names assuming format "LastName, FirstName"
    if ($entry.Name -match "^\s*(.+?),\s*(.+?)\s*$") {
        $lastName = $matches[1]
        $firstName = $matches[2]

        # Search in AD
        $user = Get-ADUser -Filter {
            Surname -eq $lastName -and GivenName -eq $firstName
        } -Properties Department

        if ($user) {
            $results += [PSCustomObject]@{
                samaccountname = $user.SamAccountName
                LastName       = $lastName
                FirstName      = $firstName
                Department     = $user.Department
            }
        } else {
            Write-Warning "User not found in AD: $firstName $lastName"
        }
    } else {
        Write-Warning "Name format invalid: $($entry.Name)"
    }
}

# Export result to CSV
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "Output written to $outputPath"

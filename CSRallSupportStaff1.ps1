# Prompt for AD credential
$username = "nih\aakangah"
$password = Read-Host "Enter password for $username" -AsSecureString
$credential = New-Object System.Management.Automation.PSCredential ($username, $password)

# Import ActiveDirectory module
Import-Module ActiveDirectory

# Prepare output array
$results = @()

# Read input file
Get-Content "C:\STPS\CSRallSupportStaff.txt" | ForEach-Object {
    $line = $_.Trim()
    if ($line -match "^(.*),\s*(.*)$") {
        $lastname_raw = $matches[1].Trim()
        $firstname_raw = $matches[2].Trim()

        # Replace "-" with space (to improve match rate)
        $lastname = $lastname_raw -replace "-", " "
        $firstname = $firstname_raw -replace "-", " "

        try {
            # Search AD using only lastname and firstname
            $user = Get-ADUser -Credential $credential -Filter {
                Surname -eq $lastname -and GivenName -eq $firstname
            } -Properties SamAccountName, GivenName, Surname, mail, department, physicalDeliveryOfficeName

            if ($user) {
                # User found
                $results += [PSCustomObject]@{
                    LastName       = $user.Surname
                    FirstName      = $user.GivenName
                    SamAccountName = $user.SamAccountName
                    Mail           = $user.mail
                    Department     = $user.department
                    Office         = $user.physicalDeliveryOfficeName
                    Status         = "Found"
                }
            }
            else {
                # User not found
                $results += [PSCustomObject]@{
                    LastName       = $lastname_raw
                    FirstName      = $firstname_raw
                    SamAccountName = ""
                    Mail           = ""
                    Department     = ""
                    Office         = ""
                    Status         = "Not Found"
                }
            }
        }
        catch {
            # If error occurs during search
            $results += [PSCustomObject]@{
                LastName       = $lastname_raw
                FirstName      = $firstname_raw
                SamAccountName = ""
                Mail           = ""
                Department     = ""
                Office         = ""
                Status         = "Error: $_"
            }
        }
    }
}

# Export everything into a single CSV file
$outputPath = "C:\STPS\CSR_SupportStaff_FullOutput.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Output "Export completed: $outputPath"

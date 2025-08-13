# Prompt for credentials
$cred = Get-Credential -UserName "nih\aakangah" -Message "Enter your NIH password"

# Define the OU to search
$searchBase = "OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"

# Prompt for user input
$firstName = Read-Host "Enter user's First Name"
$lastName = Read-Host "Enter user's Last Name"

# Import AD module
Import-Module ActiveDirectory

# Search for the user
$users = Get-ADUser -Filter {
    GivenName -eq $firstName -and Surname -eq $lastName
} -SearchBase $searchBase -Credential $cred -Properties mail, memberOf, GivenName, Surname

# Prepare output
$output = foreach ($user in $users) {
    [PSCustomObject]@{
        FirstName      = $user.GivenName
        LastName       = $user.Surname
        SAMAccountName = $user.SamAccountName
        Email          = $user.mail
        ADMemberships  = ($user.MemberOf | ForEach-Object {
            ($_ -split ',')[0] -replace '^CN='
        }) -join '; '
    }
}

# Create output directory if it doesn't exist
$outputDir = "C:\STPS"
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
}

# Export to CSV
$outputFile = Join-Path $outputDir "AD_User_Search_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$output | Export-Csv -Path $outputFile -NoTypeInformation
Write-Host "Output saved to $outputFile"

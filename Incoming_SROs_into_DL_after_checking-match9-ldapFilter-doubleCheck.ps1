# Load required modules
Import-Module ActiveDirectory
Import-Module ImportExcel

# Output path
$outputExcel = "C:\STPS\ADUser_Verification.xlsx"
$null = New-Item -Path "C:\STPS" -ItemType Directory -Force

# Prompt for credential
$credential = Get-Credential -UserName "nih\aakangah" -Message "Enter your NIH AD password"

# List of users to verify
$usersToVerify = @(
    @{ FirstName="KAUSIK";      LastName="RAY";         Email="rayk@mail.nih.gov";                   HHSID="10154608" },
    @{ FirstName="KATHERINE";   LastName="SHIM";        Email="katherine.shim@nih.gov";             HHSID="2001865974" },
    @{ FirstName="HASAN";       LastName="SIDDIQUI";    Email="hasan.siddiqui@nih.gov";             HHSID="2002165009" },
    @{ FirstName="VIKTORIYA";   LastName="SIDORENKO";   Email="viktoriya.sidorenko@nih.gov";        HHSID="2003946274" },
    @{ FirstName="SHREE";       LastName="SINGH";       Email="singhshr@mail.nih.gov";              HHSID="11608833" },
    @{ FirstName="MARISA";      LastName="SRIVAREERAT"; Email="marisa.srivareerat@nih.gov";         HHSID="2003501141" },
    @{ FirstName="ABHIGNYA";    LastName="SUBEDI";      Email="abhi.subedi@nih.gov";                HHSID="2001855030" },
    @{ FirstName="JINGSHENG";   LastName="TUO";         Email="jingsheng.tuo@nih.gov";              HHSID="10409964" }
)

# Store results
$results = @()

foreach ($user in $usersToVerify) {
    $email = $user.Email.Trim().ToLower()
    $matchedUser = $null

    try {
        # Try by mail first
        $matchedUser = Get-ADUser -Filter { mail -eq $email } -Credential $credential -Properties *
        if (-not $matchedUser) {
            # Try by UPN as fallback
            $matchedUser = Get-ADUser -Filter { userPrincipalName -eq $email } -Credential $credential -Properties *
        }

        if ($matchedUser) {
            $results += [pscustomobject]@{
                FirstName      = $user.FirstName
                LastName       = $user.LastName
                EmailSearched  = $email
                HHSID_Input    = $user.HHSID
                SamAccountName = $matchedUser.SamAccountName
                GivenName      = $matchedUser.GivenName
                Surname        = $matchedUser.Surname
                Mail           = $matchedUser.Mail
                EmployeeID     = $matchedUser.EmployeeID
                Department     = $matchedUser.Department
                FoundBy        = if ($matchedUser.Mail -eq $email) { "mail" } else { "userPrincipalName" }
            }
        } else {
            $results += [pscustomobject]@{
                FirstName      = $user.FirstName
                LastName       = $user.LastName
                EmailSearched  = $email
                HHSID_Input    = $user.HHSID
                SamAccountName = ""
                GivenName      = ""
                Surname        = ""
                Mail           = ""
                EmployeeID     = ""
                Department     = ""
                FoundBy        = "NotFound"
            }
        }
    } catch {
        Write-Warning "Error searching for $($user.FirstName) $($user.LastName): $_"
    }
}

# Export results
$results | Export-Excel -Path $outputExcel -WorksheetName "CheckResults" -AutoSize
Write-Host "`n✅ Exported verification results to: $outputExcel" -ForegroundColor Cyan

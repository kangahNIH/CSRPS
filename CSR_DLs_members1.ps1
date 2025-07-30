# Ensure the Active Directory module is available
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    try {
        Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeAllSubFeature -IncludeManagementTools -ErrorAction Stop
        Import-Module ActiveDirectory
    } catch {
        Write-Error "Active Directory module could not be installed. Please run as administrator or install RSAT manually."
        exit
    }
} else {
    Import-Module ActiveDirectory
}

# Prompt for credentials
$cred = Get-Credential -UserName "nih\aakangah" -Message "Enter your AD password"

# Define search base
$SearchBase = "OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"

# Get all mail-enabled groups
$groups = Get-ADGroup -SearchBase $SearchBase -LDAPFilter "(mail=*)" -Credential $cred -Properties Description, ManagedBy, msExchCoManagedByLink

# Initialize output array
$output = @()

foreach ($group in $groups) {
    $primaryManager = $null
    $secondaryManagers = @()

    # Resolve primary manager
    if ($group.ManagedBy) {
        try {
            $primaryManager = (Get-ADUser -Identity $group.ManagedBy -Credential $cred -Properties DisplayName).DisplayName
        } catch {
            $primaryManager = "Could not resolve"
        }
    }

    # Resolve secondary managers
    if ($group.'msExchCoManagedByLink') {
        foreach ($secMgrDN in $group.'msExchCoManagedByLink') {
            try {
                $secMgr = Get-ADUser -Identity $secMgrDN -Credential $cred -Properties DisplayName
                $secondaryManagers += $secMgr.DisplayName
            } catch {
                $secondaryManagers += "Could not resolve"
            }
        }
    }

    # Add header row for the DL
    $output += [PSCustomObject]@{
        "DL Group Name"        = $group.Name
        "Description"          = $group.Description
        "DL Manager"           = $primaryManager
        "DL Secondary Manager" = $secondaryManagers -join '; '
        "samAccountName"       = ""
        "LastName"             = ""
        "FirstName"            = ""
        "Department"           = ""
        "Office"               = ""
    }

    # Get members
    try {
        $members = Get-ADGroupMember -Identity $group.DistinguishedName -Credential $cred -Recursive | Where-Object { $_.objectClass -eq 'user' }

        $detailedMembers = foreach ($member in $members) {
            try {
                Get-ADUser -Identity $member.SamAccountName -Credential $cred -Properties sn, givenName, Department, Office |
                Select-Object @{Name="samAccountName"; Expression={$_.SamAccountName}},
                              @{Name="LastName"; Expression={$_.sn}},
                              @{Name="FirstName"; Expression={$_.givenName}},
                              Department, Office
            } catch {
                # Skip invalid member
                continue
            }
        }

        # Sort by LastName and add to output
        foreach ($dm in ($detailedMembers | Sort-Object LastName)) {
            $output += [PSCustomObject]@{
                "DL Group Name"        = ""
                "Description"          = ""
                "DL Manager"           = ""
                "DL Secondary Manager" = ""
                "samAccountName"       = $dm.samAccountName
                "LastName"             = $dm.LastName
                "FirstName"            = $dm.FirstName
                "Department"           = $dm.Department
                "Office"               = $dm.Office
            }
        }
    } catch {
        Write-Warning "Failed to retrieve members for group $($group.Name)"
    }
}

# Output to table
$output | Format-Table -AutoSize

# Optionally export to CSV
#$output | Export-Csv -Path "C:\STPS\CSR_Distribution_List_Members.csv" -NoTypeInformation -Encoding UTF8

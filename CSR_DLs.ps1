# Test by JP 8-7-2025
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

# Define target OU
$SearchBase = "OU=CSR,OU=NIH,OU=AD,DC=nih,DC=gov"

# Get all distribution groups in the OU
$groups = Get-ADGroup -SearchBase $SearchBase -LDAPFilter "(mail=*)" -Credential $cred -Properties Description, ManagedBy, msExchCoManagedByLink

# Prepare results
$results = foreach ($group in $groups) {
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

    # Resolve secondary managers (Exchange only)
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

    [PSCustomObject]@{
        "DL Group Name"      = $group.Name
        "Description"        = $group.Description
        "DL Manager"         = $primaryManager
        "DL Secondary Manager" = $secondaryManagers -join '; '
    }
}

# Output the results
$results | Format-Table -AutoSize

# Optionally export to CSV
#$results | Export-Csv -Path "C:\STPS\CSR_Distribution_Lists.csv" -NoTypeInformation -Encoding UTF8


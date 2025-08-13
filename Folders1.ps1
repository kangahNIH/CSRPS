# Prompt for Security Group Name
$securityGroupName = Read-Host "Enter Security Group Name (Example: DOMAIN\GroupName)"

# Prompt for the root folder path to search
$rootPath = Read-Host "Enter root folder path to search (Example: D:\Shares)"

# Get security group members
try {
    $group = Get-ADGroup -Identity $securityGroupName -ErrorAction Stop
    $groupMembers = Get-ADGroupMember -Identity $group | Select-Object Name, SamAccountName
} catch {
    Write-Host "Security group not found in AD." -ForegroundColor Red
    exit
}

# Create output array
$output = @()

# Search all folders recursively
Get-ChildItem -Path $rootPath -Directory -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $folderPath = $_.FullName

    # Get ACL for folder
    try {
        $acl = Get-Acl -Path $folderPath
    } catch {
        Write-Host "Cannot read ACL for $folderPath" -ForegroundColor Yellow
        return
    }

    foreach ($access in $acl.Access) {
        # Check if security group is in ACL
        if ($access.IdentityReference -eq $securityGroupName) {
            $output += [PSCustomObject]@{
                FolderPath = $folderPath
                SecurityGroup = $securityGroupName
                Permission = $access.FileSystemRights
                AccessControlType = $access.AccessControlType
                Inherited = $access.IsInherited
            }
        }
    }
}

# Output result
if ($output.Count -eq 0) {
    Write-Host "No folders found with specified security group in ACL." -ForegroundColor Yellow
} else {
    $output | Format-Table -AutoSize

    # Export result to CSV
    $csvPath = "C:\Temp\SecurityGroup_FolderPermissions.csv"
    $output | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "Results exported to $csvPath" -ForegroundColor Green
}

# Output members of security group
Write-Host "`nMembers of Security Group $securityGroupName :" -ForegroundColor Cyan
$groupMembers | Format-Table Name, SamAccountName

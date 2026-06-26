# Offboarding-WorkFlow
offboarding workflow written in powershell with an MgGraph Connection 


#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.DeviceManagement


param(
    [Parameter(Mandatory)]
    [string]$UserUPN,
    
    [Parameter(Mandatory)]
    [string]$ManagerUPN,
    
    [string]$ForwardingAddress,
    
    [int]$RetentionDays = 30,
    
    [Parameter(Mandatory)]
    [string]$TenantName,
    
    [string[]]$ExcludedGroups = @("All Company", "Everyone", "All Users"),
    [string[]]$ExcludedGroupIds = @(),
    
    [switch]$WhatIf
)

# --- Setup ---
$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$LogDir = "."
$LogPath = "$LogDir\Offboard-Log-$UserUPN-$Timestamp.json"
$AuditLog = @{ Actions = @(); Errors = @() }

# Ensure log directory exists
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# --- Connect to Microsoft Graph ---
$RequiredScopes = @(
    "User.ReadWrite.All",
    "Group.ReadWrite.All",
    "Directory.ReadWrite.All",
    "MailboxSettings.ReadWrite",
    "DeviceManagementManagedDevices.ReadWrite.All"
)
Connect-MgGraph -Scopes $RequiredScopes -NoWelcome

# --- Validate users ---
$User = Get-MgUser -UserId $UserUPN -ErrorAction SilentlyContinue
if (-not $User) { throw "User not found: $UserUPN" }

$Manager = Get-MgUser -UserId $ManagerUPN -ErrorAction SilentlyContinue
if (-not $Manager) { throw "Manager not found: $ManagerUPN" }

Write-Host "Offboarding: $($User.DisplayName) -> Manager: $($Manager.DisplayName)" -ForegroundColor Cyan

# --- Step 1: Disable account ---
Write-Host "[1/9] Disabling account..." -ForegroundColor Green
if (-not $WhatIf) {
    try {
        Update-MgUser -UserId $UserUPN -AccountEnabled:$false
        $AuditLog.Actions += "Account disabled"
    } catch { $AuditLog.Errors += "Step 1: $_" }
}

# --- Step 2: Revoke sessions ---
Write-Host "[2/9] Revoking all sessions..." -ForegroundColor Green
if (-not $WhatIf) {
    try {
        Revoke-MgUserSignInSession -UserId $UserUPN
        $AuditLog.Actions += "Sessions revoked"
    } catch { $AuditLog.Errors += "Step 2: $_" }
}

# --- Step 3: Set out-of-office ---
Write-Host "[3/9] Setting auto-reply..." -ForegroundColor Green
if (-not $WhatIf) {
    try {
        $AutoReply = @{
            Status = "Scheduled"
            ScheduledStartDateTime = @{ DateTime = (Get-Date).ToString("o"); TimeZone = "UTC" }
            ScheduledEndDateTime = @{ DateTime = (Get-Date).AddDays($RetentionDays).ToString("o"); TimeZone = "UTC" }
            ExternalReplyMessage = "Thank you for your email. I am no longer with the organization. For urgent matters, please contact $($Manager.DisplayName) at $ManagerUPN."
            InternalReplyMessage = "I have left the organization. Please contact my manager, $($Manager.DisplayName), at $ManagerUPN for related matters."
            ExternalAudience = "All"
        }
        Update-MgUserMailboxSetting -UserId $UserUPN -AutoReplySetting $AutoReply
        $AuditLog.Actions += "Auto-reply set"
    } catch { $AuditLog.Errors += "Step 3: $_" }
}

# --- Step 4: Configure mailbox forwarding ---
Write-Host "[4/9] Configuring mailbox forwarding..." -ForegroundColor Green
if (-not $WhatIf) {
    try {
        $ForwardTo = if ($ForwardingAddress) { $ForwardingAddress } else { $ManagerUPN }
        Update-MgUserMailboxSetting -UserId $UserUPN -AdditionalProperties @{
            ForwardingSmtpAddress = $ForwardTo
            DeliverToMailboxAndForward = $true
        }
        $AuditLog.Actions += "Forwarding set to $ForwardTo"
    } catch { $AuditLog.Errors += "Step 4: $_" }
}

# --- Step 5: Export and remove group memberships ---
Write-Host "[5/9] Removing group memberships..." -ForegroundColor Green
if (-not $WhatIf) {
    try {
        # Fetch all directory objects the user is a member of
        $MemberOf = Get-MgUserMemberOf -UserId $UserUPN -All

        # Filter to actual groups only, excluding dynamic groups and excluded names/IDs
        $Groups = $MemberOf | Where-Object { 
            $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group" -and
            $_.AdditionalProperties["groupTypes"] -notcontains "DynamicMembership" -and
            $_.AdditionalProperties["displayName"] -notin $ExcludedGroups -and
            $_.Id -notin $ExcludedGroupIds
        }

        $GroupExport = @()

        foreach ($g in $Groups) {
            $GroupName = $g.AdditionalProperties["displayName"]
            $GroupId   = $g.Id

            try {
                # If user is an Owner, downgrade to Member first so removal succeeds
                $Owners = Get-MgGroupOwner -GroupId $GroupId -All -ErrorAction SilentlyContinue
                $IsOwner = $Owners | Where-Object { $_.Id -eq $User.Id }
                
                if ($IsOwner) {
                    Remove-MgGroupOwnerByRef -GroupId $GroupId -DirectoryObjectId $User.Id -ErrorAction SilentlyContinue
                    Write-Host "  Removed ownership from: $GroupName" -ForegroundColor DarkGray
                }

                # Now remove from group
                Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $User.Id
                
                $GroupExport += @{
                    Id     = $GroupId
                    Name   = $GroupName
                    Status = "Removed"
                }
                Write-Host "  Removed from: $GroupName" -ForegroundColor Gray

            } catch {
                $GroupExport += @{
                    Id     = $GroupId
                    Name   = $GroupName
                    Status = "Failed: $_"
                }
                Write-Host "  Could not remove from: $GroupName (Reason: $_)" -ForegroundColor Yellow
            }
        }

        $AuditLog.Actions += "Processed $($Groups.Count) group memberships"
        $AuditLog.GroupMemberships = $GroupExport

    } catch { 
        $AuditLog.Errors += "Step 5: $_" 
    }
}

# --- Step 6: Remove licenses ---
Write-Host "[6/9] Removing licenses..." -ForegroundColor Green
if (-not $WhatIf) {
    try {
        $Licenses = Get-MgUserLicenseDetail -UserId $UserUPN
        if ($Licenses) {
            Set-MgUserLicense -UserId $UserUPN -AddLicenses @() -RemoveLicenses $Licenses.SkuId
            $AuditLog.Actions += "Removed $($Licenses.Count) license(s)"
            $AuditLog.PreviousLicenses = $Licenses | Select-Object SkuId, SkuPartNumber
        } else {
            $AuditLog.Actions += "No licenses assigned"
        }
    } catch { $AuditLog.Errors += "Step 6: $_" }
}

# --- Step 7: Retire Intune devices ---
Write-Host "[7/9] Retiring Intune devices..." -ForegroundColor Green
if (-not $WhatIf) {
    try {
        $Devices = Get-MgDeviceManagementManagedDevice -Filter "userPrincipalName eq '$UserUPN'" -ErrorAction SilentlyContinue
        if ($Devices) {
            foreach ($d in $Devices) {
                Invoke-MgRetireDeviceManagementManagedDevice -ManagedDeviceId $d.Id
                $AuditLog.Actions += "Retired device: $($d.DeviceName)"
            }
        } else {
            $AuditLog.Actions += "No Intune devices found"
        }
    } catch { $AuditLog.Errors += "Step 7: $_" }
}

# --- Step 8: OneDrive handover note ---
Write-Host "[8/9] OneDrive handover..." -ForegroundColor Green
$OneDriveUserPart = $UserUPN.Replace(".","_").Replace("@","_")
$OneDriveUrl = "https://$TenantName-my.sharepoint.com/personal/$OneDriveUserPart"
Write-Host "  Grant $ManagerUPN admin rights to: $OneDriveUrl" -ForegroundColor Yellow
$AuditLog.Actions += "OneDrive handover noted: $OneDriveUrl"

# --- Step 9: Finalize ---
Write-Host "[9/9] Finalizing..." -ForegroundColor Green
$DeletionDate = (Get-Date).AddDays($RetentionDays).ToString("yyyy-MM-dd")
$AuditLog.ScheduledDeletionDate = $DeletionDate
$AuditLog | ConvertTo-Json -Depth 5 | Out-File -FilePath $LogPath -Encoding UTF8

Write-Host "`nOffboarding complete." -ForegroundColor Cyan
Write-Host "Log: $LogPath" -ForegroundColor Gray
Write-Host "Delete after: $DeletionDate" -ForegroundColor Gray
if ($AuditLog.Errors.Count -gt 0) {
    Write-Host "Errors: $($AuditLog.Errors.Count)" -ForegroundColor Red
    $AuditLog.Errors | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}

Write-Host "`nManual next steps:" -ForegroundColor Magenta
Write-Host "1. Grant $ManagerUPN OneDrive access via https://$TenantName-admin.sharepoint.com"
Write-Host "2. Update HR/IT asset register"
Write-Host "3. Close offboarding ticket in your ticketing system"
Write-Host "4. Remove from non-Microsoft systems (VPN, GitHub, AWS, Azure DevOps, etc.)"

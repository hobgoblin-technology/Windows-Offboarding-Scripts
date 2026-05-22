# ==============================
# CONFIGURATION
# ==============================

param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$User,
    [Parameter(Mandatory=$False,Position=2)]
    [Switch]$WhatIf
)

Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited


# Resolve the exact address (handles case‑insensitivity)
$target = $User.Trim().ToLowerInvariant()


# ==============================
# FULL ACCESS PERMISSIONS
# ==============================
Write-Host "Processing FullAccess permissions..." -ForegroundColor Cyan

$Mailboxes |
Get-MailboxPermission -User $User |
Where-Object {
    $_.IsInherited -eq $false -and
    $_.AccessRights -contains "FullAccess"
} |
ForEach-Object {
    if ($WhatIf){
	Write-Host "[WHATIF] Would Remove FullAccess from $($_.Identity)" -ForegroundColor Yellow
    }
    else{
   	Write-Host "Removing FullAccess from $($_.Identity)" -ForegroundColor Yellow
    }
    
    Remove-MailboxPermission `
        -Identity $_.Identity `
        -User $User `
        -AccessRights FullAccess `
        -Confirm:$false `
        -WhatIf:$WhatIf
}

# ==============================
# SEND AS PERMISSIONS
# ==============================
Write-Host "Processing Send As permissions..." -ForegroundColor Cyan

Get-RecipientPermission -Trustee $User |
Where-Object {
    $_.AccessRights -contains "SendAs"
} |
ForEach-Object {
    if ($WhatIf){
	Write-Host "[WHATIF] Would Remove SendAs from $($_.Identity)" -ForegroundColor Yellow
    }
    else{
   	Write-Host "Removing SendAs from $($_.Identity)" -ForegroundColor Yellow
    }
    Remove-RecipientPermission `
        -Identity $_.Identity `
        -Trustee $User `
        -AccessRights SendAs `
        -Confirm:$false `
        -WhatIf:$WhatIf
}

# ==============================
# SEND ON BEHALF PERMISSIONS
# ==============================
Write-Host "Processing Send on Behalf permissions..." -ForegroundColor Cyan

$Mailboxes |
Where-Object {
    $_.GrantSendOnBehalfTo -contains $User
} |
ForEach-Object {
    if (-not $WhatIf) {
	Write-Host "Removing Send on Behalf from $($_.Identity)" -ForegroundColor Yellow
        Set-Mailbox $_.Identity `
            -GrantSendOnBehalfTo @{Remove=$User}
    } else {
        Write-Host "[WHATIF] Would remove Send on Behalf from $($_.Identity)" -ForegroundColor DarkYellow
    }
}

# ==============================
# DISTRIBUTION LIST PERMISSIONS
# ==============================
Write-Host "Processing Distribution List permissions..." -ForegroundColor Cyan

Try {
    #Get All Distribution Lists - Excluding Mail enabled security groups
    $DistributionGroups = Get-Distributiongroup -resultsize unlimited |  Where {!$_.GroupType.contains("SecurityEnabled")}
 
    #Loop through each Distribution Lists
    ForEach ($Group in $DistributionGroups)
    {
        #Check if the Distribution List contains the particular user
        If ((Get-DistributionGroupMember $Group.GUID -ResultSize Unlimited | Select -Expand PrimarySmtpAddress) -contains $User)
        {
            If (-not $WhatIf) {
                Remove-DistributionGroupMember -Identity $Group.GUID -Member $User -Confirm:$false
                Write-host "Removed user from group '$Group'" -f Green
            }
            Else{
                Write-Host -f DarkYellow "[WHAT IF] Would remove '$User' from '$Group.PrimarySmtpAddress'"
            }
        }
    }
}
Catch {
    write-host -f Red "Error:" $_.Exception.Message
}
Write-Host "Permission cleanup complete." -ForegroundColor Green

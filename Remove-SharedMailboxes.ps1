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

$Mailboxes = Get-Mailbox -ResultSize Unlimited


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
    Write-Host "Removing FullAccess from $($_.Identity)" -ForegroundColor Yellow

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
    Write-Host "Removing SendAs from $($_.Identity)" -ForegroundColor Yellow

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
    Write-Host "Removing Send on Behalf from $($_.Identity)" -ForegroundColor Yellow

    if (-not $WhatIf) {
        Set-Mailbox $_.Identity `
            -GrantSendOnBehalfTo @{Remove=$User}
    } else {
        Write-Host "WHATIF: Would remove Send on Behalf from $($_.Identity)" -ForegroundColor DarkYellow
    }
}

Write-Host "Permission cleanup complete." -ForegroundColor Green

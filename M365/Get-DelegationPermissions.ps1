<#
    .DESCRIPTION
    Queries Exchange Online mailboxes and identifies delegated access and SendAs
    permissions.

    .EXAMPLE
    Get-MailboxDelegatesReport | Export-CSV -Path .\delegates.csv
    -NoTypeInformation

    .Notes
    Requires ExchangeOnlineManagement v3 module and appropriate permissions.
#>

<#
    Derived from: https://github.com/TheTaylorLee/AdminToolbox

    Original license: MIT
    Copyright (c) Taylor Lee

    This file includes code derived from MIT-licensed source.
    See ../THIRD_PARTY_NOTICES.md for full license text and attribution.

    Modifications in this repository are distributed under The Unlicense, except
    where third-party terms apply.
#>


function Get-MailboxDelegatesReport {

    [CmdletBinding()]
    param (
    )

    # Connect to Exchange Online if not already connected
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Error "ExchangeOnlineManagement module is not installed."
        return
    }
    if (-not (Get-ConnectionInformation)) {
        Connect-ExchangeOnline -ShowBanner:$false
    }

    $mailboxes = Get-ExoMailbox -ResultSize Unlimited -Properties GrantSendOnBehalfTo

    foreach ($mbx in $mailboxes) {
        # Get Full Access Delegates
        $fullAccess = Get-EXOMailboxPermission -Identity $mbx.Identity | Where-Object {
            $_.User -notlike "NT AUTHORITY\SELF" -and $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited
        } | Select-Object -ExpandProperty User

        # Get SendAs Delegates
        $sendAs = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue | Where-Object {
            $_.Trustee -notlike "NT AUTHORITY\SELF" -and $_.AccessRights -contains "SendAs"
        } | Select-Object -ExpandProperty Trustee

        # Get Send-on-Behalf delegates (mailbox-level GrantSendOnBehalfTo).
        # The property stores a mix of display names, UPNs, and GUIDs, so resolve
        # each entry to its PrimarySmtpAddress to match the other delegate columns.
        $sendOnBehalf = foreach ($entry in $mbx.GrantSendOnBehalfTo) {
            if ([string]::IsNullOrWhiteSpace($entry)) { continue }
            $recipient = Get-Recipient -Identity $entry -ErrorAction SilentlyContinue
            if ($recipient -and $recipient.PrimarySmtpAddress) {
                $recipient.PrimarySmtpAddress
            }
            else {
                # Fall back to the raw entry if it can't be resolved
                $entry
            }
        }

        # Determine whether the account's sign-in is blocked
        $user = Get-User -Identity $mbx.Identity -ErrorAction SilentlyContinue
        if ($null -ne $user) {
            $signInBlocked = [bool]$user.BlockedCredentials
            $signInStatus = if ($signInBlocked) { 'Blocked' } else { 'Allowed' }
        }
        else {
            $signInStatus = 'Unknown'
        }

        [PSCustomObject]@{
            DisplayName         = $mbx.DisplayName
            Mailbox             = $mbx.PrimarySmtpAddress
            SignInStatus        = $signInStatus
            FullAccessDelegates = ($fullAccess -join '; ')
            SendAsDelegates     = ($sendAs -join '; ')
            SendOnBehalf        = ($sendOnBehalf -join '; ')
        }
    }
}

$reportPath = Join-Path -Path $PSScriptRoot -ChildPath "delegates.csv"
Write-Host "Generating delegation report. This may take a few minutes..."
Get-MailboxDelegatesReport | Export-Csv -Path $reportPath -NoTypeInformation
Write-Host "Delegation report exported to: $reportPath"

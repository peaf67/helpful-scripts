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

    $mailboxes = Get-ExoMailbox -ResultSize Unlimited

    foreach ($mbx in $mailboxes) {
        # Get Full Access Delegates
        $fullAccess = Get-EXOMailboxPermission -Identity $mbx.Identity | Where-Object {
            $_.User -notlike "NT AUTHORITY\SELF" -and $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited
        } | Select-Object -ExpandProperty User

        # Get SendAs Delegates
        $sendAs = Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue | Where-Object {
            $_.Trustee -notlike "NT AUTHORITY\SELF" -and $_.AccessRights -contains "SendAs"
        } | Select-Object -ExpandProperty Trustee

        [PSCustomObject]@{
            DisplayName         = $mbx.DisplayName
            Mailbox             = $mbx.PrimarySmtpAddress
            FullAccessDelegates = ($fullAccess -join '; ')
            SendAsDelegates     = ($sendAs -join '; ')
        }
    }
}

$reportPath = Join-Path -Path $PSScriptRoot -ChildPath "delegates.csv"
Write-Host "Generating delegation report. This may take a few minutes..."
Get-MailboxDelegatesReport | Export-Csv -Path $reportPath -NoTypeInformation
Write-Host "Delegation report exported to: $reportPath"

<#
    .DESCRIPTION
    Queries Exchange Online mailboxes and identifies delegated access and SendAs
    permissions.

    .EXAMPLE
    Get-MailboxDelegatesReport | Export-CSV -Path .\delegates.csv
    -NoTypeInformation

    .Notes
    Requires ExchangeOnlineManagement v3 module and appropriate permissions.
    Requires Microsoft.Graph.Users module and User.Read.All permission for
    sign-in status lookups via Microsoft Graph.

    The Graph lookup runs in a separate console process to avoid an MSAL
    assembly conflict with ExchangeOnlineManagement. A second console window
    appears briefly during the run for interactive Graph authentication.

    Tested with PowerShell 7. YMMV with PowerShell 5.
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

    # Connect to Exchange Online if not already connected. Stop the script if
    # the module is missing or authentication does not establish a connection.
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module is not installed."
    }
    if (-not (Get-ConnectionInformation)) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    if (-not (Get-ConnectionInformation)) {
        throw "Exchange Online authentication failed."
    }

    # Fetch each user's Entra ID sign-in state (accountEnabled) via Microsoft
    # Graph. Graph runs in a separate console process (a child of the current
    # PowerShell host) so its MSAL assemblies load independently, avoiding an
    # assembly-version conflict with ExchangeOnlineManagement. The child uses
    # its own window so WAM has a parent window handle for interactive auth.
    # Results are written to a temp JSON file; the exit code signals success.
    # The child is launched before the mailbox fetch so the two run concurrently.
    $childExe = if ($PSVersionTable.PSVersion.Major -ge 6) { 'pwsh.exe' } else { 'powershell.exe' }
    $childPath = Join-Path $PSHome $childExe
    $graphResultFile = [System.IO.Path]::GetTempFileName()
    $graphScriptFile = [System.IO.Path]::GetTempFileName() + '.ps1'
    $graphScript = @"
`$ErrorActionPreference = 'Stop'
try {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        Write-Error 'Microsoft.Graph.Users module is not installed.'
        exit 2
    }
    Connect-MgGraph -Scopes 'User.Read.All' | Out-Null
    if (-not (Get-MgContext)) {
        Write-Error 'Microsoft Graph authentication failed.'
        exit 3
    }
    Get-MgUser -All -Property UserPrincipalName, AccountEnabled |
        Select-Object UserPrincipalName, AccountEnabled |
        ConvertTo-Json -Depth 2 |
        Set-Content -Path '$graphResultFile' -Encoding UTF8
    exit 0
}
catch {
    Write-Error `$_.Exception.Message
    exit 4
}
"@
    Set-Content -Path $graphScriptFile -Value $graphScript -Encoding UTF8
    $graphProc = Start-Process -FilePath $childPath -ArgumentList '-NoProfile','-File',$graphScriptFile -PassThru -WindowStyle Normal

    $mailboxes = Get-ExoMailbox -ResultSize Unlimited -Properties GrantSendOnBehalfTo

    # Wait for the Graph child to finish (10-minute timeout).
    if (-not $graphProc.WaitForExit(600000)) {
        $graphProc | Stop-Process -Force -ErrorAction SilentlyContinue
        Remove-Item $graphScriptFile, $graphResultFile -Force -ErrorAction SilentlyContinue
        throw "Microsoft Graph lookup timed out."
    }
    Remove-Item $graphScriptFile -Force -ErrorAction SilentlyContinue

    if ($graphProc.ExitCode -ne 0 -or -not (Test-Path $graphResultFile)) {
        Remove-Item $graphResultFile -Force -ErrorAction SilentlyContinue
        throw "Microsoft Graph authentication or lookup failed (exit code $($graphProc.ExitCode))."
    }

    # Build a UPN -> accountEnabled hashtable from the child's JSON output.
    # Force-array the JSON parse so a single-user tenant still iterates.
    $signInByUpn = @{}
    $json = Get-Content $graphResultFile -Raw
    Remove-Item $graphResultFile -Force -ErrorAction SilentlyContinue
    if ($json) {
        foreach ($u in @($json | ConvertFrom-Json)) {
            $signInByUpn[$u.UserPrincipalName] = $u.AccountEnabled
        }
    }

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

        # Determine sign-in status from the Entra ID accountEnabled value
        # retrieved via Graph. Exchange's BlockedCredentials property does not
        # reflect actual M365 sign-in state.
        if ($signInByUpn.ContainsKey($mbx.UserPrincipalName)) {
            $enabled = $signInByUpn[$mbx.UserPrincipalName]
            $signInStatus = if ($enabled) { 'Allowed' } else { 'Blocked' }
        }
        else {
            $signInStatus = 'Unknown'
        }

        [PSCustomObject]@{
            DisplayName         = $mbx.DisplayName
            Mailbox             = $mbx.PrimarySmtpAddress
            MailboxType         = $mbx.RecipientTypeDetails
            SignInStatus        = $signInStatus
            FullAccessDelegates = ($fullAccess -join '; ')
            SendAsDelegates     = ($sendAs -join '; ')
            SendOnBehalf        = ($sendOnBehalf -join '; ')
        }
    }
}

$reportPath = Join-Path -Path $PSScriptRoot -ChildPath "delegates.csv"
Write-Host "Generating delegation report. This may take a few minutes..."
try {
    Get-MailboxDelegatesReport | Export-Csv -Path $reportPath -NoTypeInformation
}
catch {
    Write-Error "Report generation stopped: $($_.Exception.Message)"
    exit 1
}
Write-Host "Delegation report exported to: $reportPath"

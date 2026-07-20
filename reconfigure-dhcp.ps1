<#
Attribution:
Created by Joshua Pearlmutter
Repository: https://github.com/peaf67/helpful-scripts
License: Unlicense
#>

$IPType = 'IPv4'

# Optional safety defaults
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# This script requires elevation to modify adapter IP/DNS settings.
$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$currentPrincipal = [Security.Principal.WindowsPrincipal]::new($currentIdentity)
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    throw 'This script must be run in an elevated PowerShell session (Run as Administrator).'
}

# Adjust filter as needed:
# -Status Up              => only currently connected adapters
# -HardwareInterface      => skip many virtual adapters
$adapters = Get-NetAdapter | Where-Object {
    $_.Status -eq 'Up' -and $_.HardwareInterface
}

foreach ($adapter in $adapters) {
    try {
        $interface = $adapter | Get-NetIPInterface -AddressFamily $IPType -ErrorAction Stop
        $dnsConfig = $adapter | Get-DnsClientServerAddress -AddressFamily $IPType -ErrorAction Stop

        $dhcpDisabled = $interface.Dhcp -eq 'Disabled'
        $staticDnsConfigured = @($dnsConfig.ServerAddresses).Count -gt 0

        if ($dhcpDisabled -or $staticDnsConfigured) {
            $reasons = @()
            if ($dhcpDisabled) { $reasons += 'static IP' }
            if ($staticDnsConfigured) { $reasons += 'static DNS' }

            Write-Host "[$($adapter.Name)] Detected $($reasons -join ' and '). Reconfiguring..."

            if ($dhcpDisabled) {
                # Remove only IPv4 default route(s), not all routes
                Get-NetRoute -InterfaceIndex $interface.InterfaceIndex `
                             -AddressFamily IPv4 `
                             -DestinationPrefix '0.0.0.0/0' `
                             -ErrorAction SilentlyContinue |
                    Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue

                # Enable DHCP
                $interface | Set-NetIPInterface -Dhcp Enabled -ErrorAction Stop
            }

            # Reset DNS to automatic (use netsh fallback if CIM call fails)
            try {
                Set-DnsClientServerAddress -InterfaceIndex $interface.InterfaceIndex -ResetServerAddresses -ErrorAction Stop
            }
            catch {
                Write-Warning "[$($adapter.Name)] CIM DNS reset failed; trying netsh fallback. Error: $($_.Exception.Message)"
                & netsh interface ip set dns name="$($adapter.Name)" source=dhcp | Out-Null
            }

            # Give the adapter a moment, then force DHCP lease refresh
            Start-Sleep -Seconds 3
            Write-Host "[$($adapter.Name)] Releasing and renewing DHCP lease..."
            & ipconfig /release "$($adapter.Name)" | Out-Null
            Start-Sleep -Seconds 1
            & ipconfig /renew "$($adapter.Name)" | Out-Null

            Write-Host "[$($adapter.Name)] Done."
        }
        else {
            Write-Host "[$($adapter.Name)] DHCP and DNS already automatic. Skipping."
        }
    }
    catch {
        Write-Warning "[$($adapter.Name)] Failed: $($_.Exception.Message)"
    }
}

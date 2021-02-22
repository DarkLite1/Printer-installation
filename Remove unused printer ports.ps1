#Requires -Modules PrintManagement
#Requires -Version 5.1

<#
.SYNOPSIS
    Remove unused printer queues on a print server.

.DESCRIPTION
    This script is intended to run on a local server. The print ports that are 
    not in use will be removed from the system.
#>

[CmdLetBinding()]
Param ()

Process {
    Try {
        $PreInstalledPorts = @(Get-PrinterPort)
        $InstalledPrinters = @(Get-Printer)

        $PreInstalledPorts.Where( {
                ($InstalledPrinters.PortName -notcontains $_.Name) -and
                (
                    ($_.PortMonitor -eq 'Local Monitor') -or
                    ($_.PortMonitor -eq 'TCPMON.DLL')
                )
            }).foreach( {
                Remove-PrinterPort -Name $_.Name -EA Ignore
            })
    }
    Catch {
        throw "Failed removing print ports on '$env:COMPUTERNAME': $_"
    }
}
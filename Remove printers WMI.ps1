#Requires -Version 3

<#
.SYNOPSIS
    Remove printer queues on a print server that does not support the new
    printer Cmdlets of module 'PrintManagement'.

.DESCRIPTION
    This script is intended to run on a local server. The print queues will be
    removed from the system. After the desired print queues are removed, the
    ports that are unused will also be deleted.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [PSCustomObject[]]$Printers
)

Process {
    Try {
        $preInstalledPorts = Get-CimInstance -Class 'Win32_TCPIPPrinterPort'
        $preInstalledPrinters = Get-CimInstance -Class 'Win32_printer'
        $removedPrinters = @()

        #region Remove print queue
        foreach ($P in $Printers) {
            if (-not ($PrinterName = $P.PrinterName)) {
                throw "The property 'PrinterName' is mandatory"
            }

            Try {
                if (
                    $printerToRemove = $preInstalledPrinters | Where-Object {
                        $_.Name -eq $PrinterName
                    }
                ) {
                    $printerToRemove | Remove-CimInstance -ErrorAction Stop

                    $P.Action += 'Removed printer queue'
                    $removedPrinters += $printerToRemove
                }
                else {
                    $P.Error += 'Printer queue not found'
                    Continue
                }
            }
            Catch {
                $P.Error += "Error: $_"
            }
        }
        #endregion

        #region Remove printer ports
        $PostInstalledPrinters = @(Get-CimInstance -ClassName 'Win32_printer' -Namespace 'root\CIMV2')
        $PortsToRemove = @{ }

        Try {
            foreach ($P in @($removedPrinters.PortName | Sort-Object -Unique)) {
                $PortToRemove = $preInstalledPorts | Where-Object { $_.Name -eq $P }

                $PrePrintersWithPort = $preInstalledPrinters | Where-Object {
                    $_.PortName -eq $P
                } |
                Select-Object -ExpandProperty 'Name'

                $PostPrintersWithPort = $PostInstalledPrinters | Where-Object {
                    $_.PortName -eq $P
                } |
                Select-Object -ExpandProperty 'Name'

                $PortsToRemove.$P = @{
                    PrePrintersWithPort  = $PrePrintersWithPort
                    PostPrintersWithPort = $PostPrintersWithPort
                    Action               = $null
                    Error                = $null
                }

                if ($PostPrintersWithPort) {
                    $PortsToRemove.$P.Error = "Port '$P > $($PortToRemove.HostAddress)' cannot be removed because it is in use by {0} '{1}'" -f $(if ($PostPrintersWithPort -eq 1) { 'printer' }else { 'printers' }), $($PostPrintersWithPort -join "', '")
                }
                else {
                    Try {
                        $PortToRemove | Remove-CimInstance -ErrorAction Stop

                        $PortsToRemove.$P.Action = "Removed port '$P > $($PortToRemove.HostAddress)'"
                    }
                    Catch {
                        $PortsToRemove.$P.Error = $_
                    }
                }
            }

            foreach ($P in $PortsToRemove.GetEnumerator()) {
                $Printers.Where(
                    { $P.Value.PrePrintersWithPort -contains $_.PrinterName }
                ).foreach(
                    {
                        if ($P.Value.Error) {
                            $_.Error += $P.Value.Error
                        }
                        if ($P.Value.Action) {
                            $_.Action += $P.Value.Action
                        }
                    }
                )
            }
        }
        Catch {
            Throw "Failed removing printer ports: $_"
        }
        #endregion

        #region Set status
        Try {
            foreach ($P in $Printers) {
                if ($P.Error) {
                    $P.Status = 'Error'
                    Write-Warning "'$($P.PrinterName)' Error '$($P.Error -join ', ')'"
                }
                elseif (
                    ($preInstalledPrinters.Name -contains $P.PrinterName) -and
                    ($PostInstalledPrinters.Name -notContains $P.PrinterName)
                ) {
                    $P.Status = 'Removed'
                }
                else {
                    $P.Status = 'Error'
                }
            }
        }
        Catch {
            Throw "Failed setting status: $_"
        }
        #endregion

        $Printers
    }
    Catch {
        throw "Failed removing print queues on '$env:COMPUTERNAME': $_"
    }
}
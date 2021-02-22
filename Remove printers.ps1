#Requires -Modules PrintManagement
#Requires -Version 5.1

<#
.SYNOPSIS
    Remove printer queues on a print server.

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

Begin {
    Function Remove-PrinterPortHC {
        <#
        .SYNOPSIS
            Remove a printer port.

        .DESCRIPTION
            In some cases when a printer has been removed recently it seems 
            to might take a while until the printer is really removed from 
            the system and the port can be properly removed without the 
            error that a printer is still using it.

        .EXAMPLE
            Remove-PrinterPortHC -Name '10.10.10.10'

            Will remove the port that is named '10.10.10.10' and when it fails 
            will keep on trying to remove it for 17 seconds.
        #>

        Param (
            [String]$Name,
            [Int]$RetryTimes = 17
        )

        $R = 0
        $Done = $false

        while (($R -le $RetryTimes) -and (-not $Done)) {
            Try {
                Get-PrinterPort -Name $Name -EA Ignore | Remove-PrinterPort -EA Stop
                $Done = $true
            }
            Catch {
                $R += 1
                Start-Sleep -Seconds 1
            }
        }

        if (-not $Done) {
            throw $Error[0]
        }
    }
}

Process {
    Try {
        $PreInstalledPorts = @(Get-PrinterPort)
        $PreInstalledPrinters = @(Get-Printer)
        $RemovedPrinters = @()

        #region Remove print queue
        foreach ($P in $Printers) {
            if (-not ($PrinterName = $P.PrinterName)) {
                throw "The property 'PrinterName' is mandatory"
            }

            Try {
                if ($PrinterToRemove = $PreInstalledPrinters.Where( { $_.Name -eq $PrinterName })) {
                    Remove-Printer -Name $PrinterToRemove.Name
                    $P.Action += 'Removed printer queue'
                    $RemovedPrinters += $PrinterToRemove
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
        $PostInstalledPrinters = @(Get-Printer)
        $PortsToRemove = @{ }

        Try {
            foreach ($P in @($RemovedPrinters.PortName | Sort-Object -Unique)) {
                $PrinterHostAddress = $PreInstalledPorts.Where( { $_.Name -eq $P }) |
                Select-Object -First 1 -ExpandProperty 'PrinterHostAddress'

                $PrePrintersWithPort = $PreInstalledPrinters.Where( { $_.PortName -eq $P }) |
                Select-Object -ExpandProperty 'Name'

                $PostPrintersWithPort = $PostInstalledPrinters.Where( { $_.PortName -eq $P }) |
                Select-Object -ExpandProperty 'Name'

                $PortsToRemove.$P = @{
                    PrePrintersWithPort  = $PrePrintersWithPort
                    PostPrintersWithPort = $PostPrintersWithPort
                    Action               = $null
                    Error                = $null
                }

                if ($PostPrintersWithPort) {
                    $PortsToRemove.$P.Error = "Port '$P > $PrinterHostAddress' cannot be removed because it is in use by {0} '{1}'" -f $(if ($PostPrintersWithPort -eq 1) { 'printer' }else { 'printers' }), $($PostPrintersWithPort -join "', '")
                }
                else {
                    Try {
                        Remove-PrinterPortHC -Name $P
                        $PortsToRemove.$P.Action = "Removed port '$P > $PrinterHostAddress'"
                    }
                    Catch {
                        $PortsToRemove.$P.Error = $_
                    }
                }
            }

            foreach ($P in $PortsToRemove.GetEnumerator()) {
                $Printers.Where( { $P.Value.PrePrintersWithPort -contains $_.PrinterName }).foreach( {
                        if ($P.Value.Error) {
                            $_.Error += $P.Value.Error
                        }
                        if ($P.Value.Action) {
                            $_.Action += $P.Value.Action
                        }
                    })
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
                    ($PreInstalledPrinters.Name -contains $P.PrinterName) -and
                    ($PostInstalledPrinters.Name -notcontains $P.PrinterName)
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
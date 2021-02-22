#Requires -Modules PrintManagement
#Requires -Version 5.1

<#
.SYNOPSIS
    Install printer queues on a print server.

.DESCRIPTION
    This script is intended to run on a local server. The print queues will be 
    created and/or corrected to the desired configuration.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [PSCustomObject[]]$Printers
)

Begin {
    Function Get-IncorrectConfigHC {
        <#
        .SYNOPSIS
            Retrieve property names that are not the same.

        .DESCRIPTION
            Test the printer configuration settings to verify which properties 
            are not the same. Compare the hashtable coming from the import file 
            with the actual printer configuration and return a hashtable with 
            all incorrect keys and values. So they can be set with 
            Set-PrinterConfiguration.

            Values in the ReferenceObject that are NULL are simply ignored. In 
            this case it is assumed that the default printer settings or 
            whatever is there is ok with the requester.

        .EXAMPLE
            $Request = @{
                Collate       = $true
                Color         = $true
                DuplexingMode = 'TwoSidedLongEdge'
                PaperSize     = 'A4'
                PrinterName   = $testPrinter.PrinterName
            }

            $ConfigParams = @{
                Collate       = $true
                Color         = $false
                DuplexingMode = 'TwoSidedLongEdge'
                PaperSize     = 'A4'
                PrinterName   = $testPrinter.PrinterName
            }
            Set-PrintConfiguration @ConfigParams

            $PrintConfig = Get-PrintConfiguration $testPrinter.PrinterName

            Get-IncorrectConfigHC -ReferenceObject $Request -DifferenceObject $PrintConfig -Property @(
                'Published', 'Shared', 'ShareName', 'KeepPrintedJobs', 'Location', 'Comment'
            )

            OUTPUT
            Returns the property 'Color' as being incorrect.
            #>

        [OutputType([HashTable])]
        Param (
            [Parameter(Mandatory)]
            [PSCustomObject]$ReferenceObject,
            [Parameter(Mandatory)]
            [Array]$DifferenceObject,
            [Parameter(Mandatory)]
            [String[]]$Property,
            [Switch]$TestNullValue
        )

        $Result = @{ }

        if ($TestNullValue) {
            @($ReferenceObject.PSObject.Properties).where( {
                    ($Property.Contains($_.Name)) -and
                    ($_.Value -ne $DifferenceObject.($_.Name))
                }).Foreach( {
                    $Result[$_.Name] = $_.Value
                })
        }
        else {
            @($ReferenceObject.PSObject.Properties).where( {
                    ($Property.Contains($_.Name)) -and
                    ($_.Value -ne $DifferenceObject.($_.Name)) -and
                    ($null -ne $_.Value)
                }).Foreach( {
                    $Result[$_.Name] = $_.Value
                })
        }

        if ($Result.Count -ne 0) {
            $Result
        }
    }

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

            Will remove the port that is named '10.10.10.10' and when it 
            fails will keep on trying to remove it for 17 seconds.
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
        $InitiallyInstalledPrinters = @(Get-Printer)

        #region Add missing drivers
        $InstalledDrivers = Get-PrinterDriver

        $Printers | Group-Object DriverName |
        Where-Object { -not ($InstalledDrivers.Name -contains $_.Name) } | ForEach-Object {
            Try {
                $DriverName = $_.Name
                $Group = @($_.Group)

                Add-PrinterDriver -Name $DriverName -EA Stop

                $Group.foreach( {
                        $_.Action += 'Installed driver'
                    })
            }
            Catch {
                $Group.foreach( {
                        $_.Error += 'Driver not available'
                    })
            }
        }
        #endregion

        #region Printer ports
        Write-Verbose 'Ports'

        $InstalledPorts = @(Get-PrinterPort)

        foreach ($P in ($Printers.Where( { -not $_.Error }))) {
            Try {
                $PortName = $P.PortName
                $PortHostAddress = $P.PortHostAddress

                $InstalledPort = $InstalledPorts.Where( { $_.Name -eq $PortName })

                if (-not $InstalledPort) {
                    #region Install missing printer port
                    Add-PrinterPort -Name $PortName -PrinterHostAddress $PortHostAddress

                    $P.Action += 'Installed printer port'
                    $InstalledPorts = @(Get-PrinterPort)
                    #endregion
                }
                elseif ($InstalledPort.PrinterHostAddress -ne $PortHostAddress) {

                    $InstalledPrintersWithPort = $InitiallyInstalledPrinters.Where( {
                            $_.PortName -eq $PortName })

                    if (-not $InstalledPrintersWithPort) {
                        #region Correct port that is not in use
                        Remove-PrinterPortHC -Name $PortName
                        Add-PrinterPort -Name $PortName -PrinterHostAddress $PortHostAddress

                        $P.Action += 'Corrected printer port'
                        $InstalledPorts = @(Get-PrinterPort)
                        #endregion
                    }
                    else {
                        $PrintersWithPort = $Printers.Where( { $_.PortName -eq $PortName })

                        if ($InstalledPrintersWithPort.Where( {
                                    -not ($PrintersWithPort.PrinterName -contains $_.Name)
                                })) {
                            $P.Error += "Port in use by '$($InstalledPrintersWithPort.Name)'"
                        }
                        else {
                            #region Correct port that is in use only by printers in the import file
                            $InstalledPrintersWithPort.foreach( {
                                    Remove-Printer -Name $_.Name -ErrorAction Stop
                                })

                            Remove-PrinterPortHC -Name $PortName
                            Add-PrinterPort -Name $PortName -PrinterHostAddress $PortHostAddress

                            $P.Action += 'Corrected printer port'
                            $InstalledPorts = @(Get-PrinterPort)
                            #endregion
                        }
                    }
                }
            }
            Catch {
                $P.Error += "Error: $_"
            }

        }
        #endregion

        #region Printers
        $InstalledPorts = @(Get-PrinterPort)
        $InstalledPrinters = @(Get-Printer)

        foreach ($P in ($Printers.Where( { -not $_.Error }))) {
            Try {
                $PrinterName = $P.PrinterName
                $DriverName = $P.DriverName

                $InstalledPrinter = $InstalledPrinters.Where( { $_.Name -eq $PrinterName }, 'First', 1)

                #region Add missing printer and correct driver
                if (-not $InstalledPrinter) {
                    $AddPrinterParams = @{
                        Name       = $PrinterName
                        DriverName = $DriverName
                        PortName   = $P.PortName
                    }
                    Add-Printer @AddPrinterParams
                    $P.Action += 'Added printer'

                    $InstalledPrinters = @(Get-Printer)
                }
                elseif ($InstalledPrinter.DriverName -ne $DriverName) {
                    Set-Printer -Name $PrinterName -DriverName $DriverName
                    $P.Action += 'Corrected driver'

                    $InstalledPrinters = @(Get-Printer)
                }
                #endregion

                #region Set printconfig
                $PrintConfig = Get-PrintConfiguration -PrinterName $PrinterName

                $GetIncorrectConfParams = @{
                    ReferenceObject  = $P
                    DifferenceObject = $PrintConfig
                    Property         = @('Collate', 'Color', 'DuplexingMode', 'PaperSize')
                }
                if ($IncorrectConfig = Get-IncorrectConfigHC @GetIncorrectConfParams) {
                    Set-PrintConfiguration @IncorrectConfig -PrinterName $PrinterName

                    $IncorrectConfig.GetEnumerator().ForEach( {
                            $P.Action += "Set $($_.Key)"
                        })
                }

                $PrintConfig = Get-Printer -Name $PrinterName

                $GetIncorrectConfParams = @{
                    ReferenceObject  = $P
                    DifferenceObject = $PrintConfig
                    Property         = @('Published', 'Shared', 'ShareName', 'KeepPrintedJobs')
                }
                if ($IncorrectConfig = Get-IncorrectConfigHC @GetIncorrectConfParams) {
                    Set-Printer @IncorrectConfig -Name $PrinterName

                    $IncorrectConfig.GetEnumerator().ForEach( {
                            $P.Action += "Set $($_.Key)"
                        })
                }

                if (-not $PrintConfig.Comment) { $PrintConfig.Comment = $null }
                if (-not $PrintConfig.Location) { $PrintConfig.Location = $null }

                $GetIncorrectConfParams = @{
                    ReferenceObject  = $P
                    DifferenceObject = $PrintConfig
                    Property         = @('Location', 'Comment')
                    TestNullValue    = $true
                }
                if ($IncorrectConfig = Get-IncorrectConfigHC @GetIncorrectConfParams) {
                    Set-Printer @IncorrectConfig -Name $PrinterName

                    $IncorrectConfig.GetEnumerator().ForEach( {
                            $P.Action += "Set $($_.Key)"
                        })
                }
                #endregion
            }
            Catch {
                $P.Error += "Error: $_"
            }
        }
        #endregion

        #region Set status
        $PostInstalledPrinters = @(Get-Printer)

        foreach ($P in $Printers) {
            if ($P.Error) {
                $P.Status = 'Error'
                Write-Warning "'$($P.PrinterName)' Error '$($P.Error -join ', ')'"
            }
            elseif ($PostInstalledPrinters.Name -contains $P.PrinterName) {
                if ($InitiallyInstalledPrinters.Name -notcontains $P.PrinterName) {
                    $P.Status = 'Installed'
                }
                else {
                    if ($P.Action) {
                        $P.Status = 'Updated'
                    }
                    else {
                        $P.Status = 'Ok'
                    }
                }
            }
            else {
                $P.Status = 'Error'
            }
        }
        #endregion

        $Printers
    }
    Catch {
        throw "Failed configuring print server '$env:COMPUTERNAME': $_"
    }
}
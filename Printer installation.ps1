#Requires -Modules PrintManagement, ImportExcel
#Requires -Version 5.1

<#
.SYNOPSIS
    Install printer queues on a print server.

.DESCRIPTION
    An Excel file filled with one row per printer is used as input for the 
    script. These print queues will be created on the print server with their 
    defined settings.

    Every printer queue is verified to see if the settings match with the 
    requested settings in the Excel file. When that's not the case The settings 
    will be overwritten/corrected.

    The correcting of printer queues will only happen when the printer port is 
    not in use by printers not mentioned in the Excel file. We do this to avoid 
    breaking existing printer queues that are not in the file.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$ScriptTestComputer = 'Test computer compliant.ps1',
    [String]$ScriptAddPrinters = 'Add printers.ps1',
    [String]$ScriptRemovePrinters = 'Remove printers.ps1',
    [String]$ScriptRemovePrintersWMI = 'Remove printers WMI.ps1',
    [String]$ScriptRemoveUnusedPorts = 'Remove unused printer ports.ps1',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Printers\Printer installation\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Function Get-IncorrectConfigHC {
        <#
        .SYNOPSIS
            Retrieve property names that are not the same.

        .DESCRIPTION
            Test the printer configuration settings to verify which 
            properties are are not the same. Compare the hashtable coming from 
            the import file with the actual printer configuration and return a 
            hashtable with all incorrect keys and values. So they can be set 
            with Set-PrinterConfiguration.

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

            Get-IncorrectConfigHC -ReferenceObject $Request -DifferenceObject $PrintConfig

            OUTPUT
            Returns the property 'Color' as being incorrect.
        #>

        [OutputType([HashTable])]
        Param (
            [Parameter(Mandatory)]
            [PSCustomObject]$ReferenceObject,
            [Parameter(Mandatory)]
            [Array]$DifferenceObject,
            [String[]]$Property = @(
                'Collate',
                'Color',
                'DuplexingMode',
                'PaperSize'
            )
        )

        $Result = @{ }

        @($ReferenceObject.PSObject.Properties).where( {
                ($Property.Contains($_.Name)) -and
                ($_.Value -ne $DifferenceObject.($_.Name)) -and
                ($null -ne $_.Value)
            }).Foreach( {
                $Result[$_.Name] = $_.Value
            })

        if ($Result) {
            $Result
        }
    }
    
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start
        $Error.Clear()

        if (-not (Test-Path -Path $ImportFile -PathType Leaf)) {
            throw "Import file '$ImportFile' not found"
        }

        #region Test 'Test computer' script
        Try {
            $ScriptTestComputerItem = Get-Item (Join-Path -Path $PSScriptRoot -ChildPath $ScriptTestComputer) -EA Stop
        }
        Catch {
            throw "The test computer script file '$ScriptTestComputer' was not found"
        }
        #endregion

        #region Test 'Add printers' script
        Try {
            $ScriptAddPrintersItem = Get-Item (Join-Path -Path $PSScriptRoot -ChildPath $ScriptAddPrinters) -EA Stop
        }
        Catch {
            throw "The add printers script file '$ScriptAddPrinters' was not found"
        }
        #endregion

        #region Test 'Remove printers' script
        Try {
            $ScriptRemovePrintersItem = Get-Item (Join-Path -Path $PSScriptRoot -ChildPath $ScriptRemovePrinters) -EA Stop
        }
        Catch {
            throw "The remove printers script file '$ScriptRemovePrinters' was not found"
        }
        #endregion

        #region Test 'Remove printers WMI' script
        Try {
            $ScriptRemovePrintersWMIItem = Get-Item (Join-Path -Path $PSScriptRoot -ChildPath $ScriptRemovePrintersWMI) -EA Stop
        }
        Catch {
            throw "The remove printers WMI script file '$ScriptRemovePrintersWMI' was not found"
        }
        #endregion

        #region Test 'Remove unused printer ports' script
        Try {
            $ScriptRemoveUnusedPortsItem = Get-Item (Join-Path -Path $PSScriptRoot -ChildPath $ScriptRemoveUnusedPorts) -EA Stop
        }
        Catch {
            throw "The remove unused printer ports script file '$ScriptRemoveUnusedPorts' was not found"
        }
        #endregion
        
        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import file properties
        $PrintersWorksheetProperties = @{
            Color           = @{
                Type      = 'Boolean'
                Mandatory = $false
                DenySpace = $null
                DenyDots  = $null
            }
            Collate         = @{
                Type      = 'Boolean'
                Mandatory = $false
                DenySpace = $null
                DenyDots  = $null
            }
            Comment         = @{
                Type      = 'String'
                Mandatory = $false
                DenySpace = $false
                DenyDots  = $false
            }
            DuplexingMode   = @{
                Type      = 'String'
                Mandatory = $false
                DenySpace = $null
                DenyDots  = $null
            }
            KeepPrintedJobs = @{
                Type      = 'Boolean'
                Mandatory = $false
                DenySpace = $null
                DenyDots  = $null
            }
            Location        = @{
                Type      = 'String'
                Mandatory = $false
                DenySpace = $false
                DenyDots  = $false
            }
            Published       = @{
                Type      = 'Boolean'
                Mandatory = $false
                DenySpace = $null
                DenyDots  = $null
            }
            DriverName      = @{
                Type      = 'String'
                Mandatory = $true
                DenySpace = $false
                DenyDots  = $false
            }
            PaperSize       = @{
                Type      = 'String'
                Mandatory = $false
                DenySpace = $null
                DenyDots  = $null
            }
            PortHostAddress = @{
                Type      = 'String'
                Mandatory = $true
                DenySpace = $true
                DenyDots  = $false
            }
            PortName        = @{
                Type      = 'String'
                Mandatory = $true
                DenySpace = $true
                DenyDots  = $false
            }
            PrinterName     = @{
                Type      = 'String'
                Mandatory = $true
                DenySpace = $true
                DenyDots  = $true
            }
            ServerName      = @{
                Type      = 'String'
                Mandatory = $true
                DenySpace = $true
                DenyDots  = $true
            }
            ShareName       = @{
                Type      = 'String'
                Mandatory = $false
                DenySpace = $true
                DenyDots  = $true
            }
        }

        $RemoveWorksheetProperties = @{
            ServerName  = @{
                Type      = 'String'
                Mandatory = $true
                DenySpace = $true
                DenyDots  = $true
            }
            PrinterName = @{
                Type      = 'String'
                Mandatory = $true
                DenySpace = $true
                DenyDots  = $true
            }
        }

        $DuplexingModeEnum = @('OneSided', 'TwoSidedLongEdge', 'TwoSidedShortEdge')

        $PaperSizeEnum = [Enum]::GetNames([System.Drawing.Printing.PaperKind])
        #endregion

        #region Import worksheets
        $RemoveWorksheetErrors = $PrintersWorksheetErrors = $ConflictingWorksheetErrors = @()

        Write-EventLog @EventVerboseParams -Message "Import Excel sheet '$ImportFile'"

        $PrintersWorksheet = @(Import-Excel -Path $ImportFile -WorksheetName 'Printers' |
            Remove-ImportExcelHeaderProblemOnEmptySheetHC |
            Select-Object -Property @{N = 'Status'; E = { $null } }, *,
            @{N = 'Action'; E = { , @() } },
            @{N = 'Error'; E = { $null } } -ExcludeProperty Status, Action, Error)

        $RemoveWorksheet = @(Import-Excel -Path $ImportFile -WorksheetName 'Remove' |
            Remove-ImportExcelHeaderProblemOnEmptySheetHC |
            Select-Object -Property @{N = 'Status'; E = { $null } }, *,
            @{N = 'Action'; E = { , @() } },
            @{N = 'Error'; E = { , @() } } -ExcludeProperty Status, Action, Error )
        #endregion

        #region Test worksheet Printers
        Write-EventLog @EventVerboseParams -Message "Test worksheet 'Printers'"

        foreach ($P in $PrintersWorksheet) {
            #region Remove leading and trailing spaces
            @($PrintersWorksheetProperties.GetEnumerator().where( {
                        ($_.Value.Type -eq 'String') -and
                        ($P.PSObject.Properties.Name -contains $_.Name  )
                    })).ForEach( {
                    $P.($_.Name) = if (
                        ($P.($_.Name)) -and ($tmp = $P.($_.Name).Trim())) {
                        $tmp
                    }
                    else {
                        $null
                    }
                })
            #endregion

            #region Test mandatory properties
            @($PrintersWorksheetProperties.GetEnumerator().where( { $_.Value.Mandatory }).Name).Where( { -not ($P.$_) }).ForEach( {
                    $PrintersWorksheetErrors += "Printer '$($P.PrinterName)' on server '$($P.ServerName)' is missing property '$_'"
                })
            #endregion

            #region Test boolean properties
            @($PrintersWorksheetProperties.GetEnumerator().where( { $_.Value.Type -eq 'Boolean' }).Name).where( {
                    ($P.$_) -and (-not ($P.$_ -is [Boolean]))
                }).ForEach( {
                    $PrintersWorksheetErrors += "For printer '$($P.PrinterName)' boolean property '$_' with value '$($P.$_)' is not valid. Only TRUE, FALSE or NULL are supported"
                })
            #endregion

            #region Test DuplexingMode value
            if (($P.DuplexingMode) -and (-not ($DuplexingModeEnum -contains $P.DuplexingMode))) {
                $PrintersWorksheetErrors += "For printer '$($P.PrinterName)' the property DuplexingMode has value '$($P.DuplexingMode)' which is not valid. Supported values are blank or $($DuplexingModeEnum.Where({$_}) -join ', ')'."
            }
            #endregion

            #region Test PaperSize value
            if (($P.PaperSize) -and (-not ($PaperSizeEnum -contains $P.PaperSize))) {
                $PrintersWorksheetErrors += "For printer '$($P.PrinterName)' the property PaperSize has value '$($P.PaperSize)' which is not valid. Supported values are blank or $($PaperSizeEnum.Where({$_}) -join ', ')'."
            }
            #endregion

            #region Test for unknown properties
            @($P.PSObject.Properties.Name).Where( {
                    ($PrintersWorksheetProperties.Keys -notcontains $_ ) -and
                    (@('Status', 'Action', 'Error') -notcontains $_)
                }).foreach( {
                    $PrintersWorksheetErrors += "Property '$_' is not valid, only the following prperties are allowed: '$($PrintersWorksheetProperties.Keys -join ', ')'."
                })
            #endregion

            #region Test properties cannot have spaces
            @($PrintersWorksheetProperties.GetEnumerator().where( { $_.Value.DenySpace }).Name).Where( { $P.$_ -match '\s' }).ForEach( {
                    $PrintersWorksheetErrors += "The field '$_' with value '$($P.$_)' cannot contain spaces."
                })
            #endregion

            #region Test properties cannot have dots
            @($PrintersWorksheetProperties.GetEnumerator().where( { $_.Value.DenyDots }).Name).Where( { $P.$_ -match '\.' }).ForEach( {
                    $PrintersWorksheetErrors += "The field '$_' with value '$($P.$_)' cannot contain dots."
                })
            #endregion
        }

        #region Duplicate PrinterName on the same server
        $PrintersWorksheet | Group-Object ServerName, PrinterName | Where-Object { $_.Count -GE 2 } |
        ForEach-Object {
            $PrintersWorksheetErrors += "Duplicate PrinterName, ServerName combination: $($_.Name)"
        }
        #endregion

        #region Conflicting PortHostAddress for a single PortName on the same server
        foreach ($S in ($PrintersWorksheet | Group-Object ServerName)) {
            @($S.Group | Group-Object PortName | Where-Object { $_.Count -GE 2 }).where( {
                    @($_.Group | Select-Object PortHostAddress -Unique).Count -GE 2 }).ForEach( {
                    $_.Group.foreach( {
                            $PrintersWorksheetErrors += "Printer '$($_.PrinterName)' on server '$($_.ServerName)' has a conflicting PortHostAddress '$($_.PortHostAddress)' for PortName '$($_.PortName)'"
                        })
                })
        }
        #endregion
        #endregion

        #region Test Worksheet Remove
        Write-EventLog @EventVerboseParams -Message "Test worksheet 'Remove'"

        foreach ($R in $RemoveWorksheet) {
            #region Remove leading and trailing spaces
            @($RemoveWorksheetProperties.GetEnumerator().where( {
                        ($_.Value.Type -eq 'String') -and
                        ($R.PSObject.Properties.Name -contains $_.Name  )
                    })).ForEach( {
                    $R.($_.Name) = if (
                        ($R.($_.Name)) -and ($tmp = $R.($_.Name).Trim())) {
                        $tmp
                    }
                    else {
                        $null
                    }
                })
            #endregion

            #region Test mandatory properties
            @($RemoveWorksheetProperties.GetEnumerator().where( { $_.Value.Mandatory }).Name).Where( { -not ($R.$_) }).ForEach( {
                    $RemoveWorksheetErrors += "The mandatory property '$_' is missing."
                })
            #endregion

            #region Test properties cannot have spaces
            @($RemoveWorksheetProperties.GetEnumerator().where( { $_.Value.DenySpace }).Name).Where( { $R.$_ -match '\s' }).ForEach( {
                    $RemoveWorksheetErrors += "The field '$_' with value '$($R.$_)' cannot contain spaces."
                })
            #endregion

            #region Test properties cannot have dots
            @($RemoveWorksheetProperties.GetEnumerator().where( { $_.Value.DenyDots }).Name).Where( { $R.$_ -match '\.' }).ForEach( {
                    $RemoveWorksheetErrors += "The field '$_' with value '$($R.$_)' cannot contain dots."
                })
            #endregion
        }

        #region Duplicate PrinterName on the same server
        $RemoveWorksheet | Group-Object ServerName, PrinterName | Where-Object { $_.Count -GE 2 } |
        ForEach-Object {
            $RemoveWorksheetErrors += "Duplicate PrinterName, ServerName combination: $($_.Name)"
        }
        #endregion
        #endregion

        #region Test duplicate records between the worksheet Printers and Remove
        $Template = "ServerName '{0}' PrinterName '{1}'"

        $PrintersWorksheetCombos = $PrintersWorksheet.Where( { $_.ServerName -and $_.PrinterName }).foreach( {
                $Template -f $_.ServerName, $_.PrinterName
            })

        $RemoveWorksheet.foreach( {
                $RemoveWorksheetCombo = ($Template -f $_.ServerName, $_.PrinterName)
                if ($DuplicateCombo = $PrintersWorksheetCombos -contains $RemoveWorksheetCombo) {
                    $ConflictingWorksheetErrors += $DuplicateCombo
                }
            })
        #endregion

        #region Send mail to users on incorrect input
        if ($PrintersWorksheetErrors -or $RemoveWorksheetErrors -or $ConflictingWorksheetErrors) {
            $MailParams = @{
                LogFolder = $LogParams.LogFolder
                Header    = $ScriptName
                Save      = $LogFile + ' - Mail.html'
                To        = $MailTo
                Bcc       = $ScriptAdmin
                Message   = "<p>Incorrect data found in the Excel import file.</p>"
                Subject   = 'FAILURE - Incorrect input'
                Priority  = 'High'
            }

            if ($PrintersWorksheetErrors = $PrintersWorksheetErrors | Sort-Object -Unique) {
                $MailParams.Message += ("<p>Worksheet '<b>Printers</b>':</p>" + ($PrintersWorksheetErrors | ConvertTo-HtmlListHC))

                $WarningMessage = "Worksheet 'Printers':`n`n- $($PrintersWorksheetErrors -join "`n")"
                Write-EventLog @EventErrorParams -Message $WarningMessage

                <# To avoid writing strings to the event log that are to large:
                Write-EventLog @EventErrorParams -Message (
                    $WarningMessage.Substring(0, [System.Math]::Min($WarningMessage.Length, 3000))
                )
                #>
                Write-Warning $WarningMessage
            }
            if ($RemoveWorksheetErrors = $RemoveWorksheetErrors | Sort-Object -Unique) {
                $MailParams.Message += ("<p>Worksheet '<b>Remove</b>':</p>" + ($RemoveWorksheetErrors | ConvertTo-HtmlListHC))

                $WarningMessage = "Worksheet 'Remove':`n`n- $($RemoveWorksheetErrors -join "`n")"
                Write-EventLog @EventErrorParams -Message $WarningMessage
                Write-Warning $WarningMessage
            }
            if ($ConflictingWorksheetErrors = $ConflictingWorksheetErrors | Sort-Object -Unique) {
                $MailParams.Message += ("<p>Worksheet '<b>Printers</b>' and '<b>Remove</b>' contain duplicate records:</p>" + ($ConflictingWorksheetErrors | ConvertTo-HtmlListHC))

                $WarningMessage = "Worksheet 'Printers' and 'Remove' contain duplicate records:`n`n- $($ConflictingWorksheetErrors -join "`n")"
                Write-EventLog @EventErrorParams -Message $WarningMessage
                Write-Warning $WarningMessage
            }
            Send-MailHC @MailParams

            Write-EventLog @EventEndParams; Exit
        }

        Write-EventLog @EventVerboseParams -Message 'All tests on the import file passed'
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_  -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        foreach ($P in $PrintersWorksheet) {
            #region Set Shared value
            try {
                $P | Add-Member -NotePropertyName 'Shared' -NotePropertyValue $false -EA Stop
            }
            catch {
                throw "The property 'Shared' is not allowed, please use the property 'ShareName' instead. When the 'ShareName' value is blank, 'Shared' is set to FALSE. When 'ShareName' has a value, 'Shared' is set to TRUE."
            }

            $P.Shared = if ($P.ShareName) { $true }
            #endregion
        }

        #region Call the 'Test computer' script
        $UniqueComputerNames = @($RemoveWorksheet + $PrintersWorksheet) |
        Group-Object ServerName -NoElement | Select-Object -ExpandProperty Name

        [Array]$CompliantComputers = if ($UniqueComputerNames) {
            Write-EventLog @EventOutParams -Message "Start 'Test computer' script on:`n`n$($UniqueComputerNames -join "`n")"

            Try {
                $InvokeParams = @{
                    FilePath     = $ScriptTestComputerItem
                    ComputerName = $UniqueComputerNames
                }
                Invoke-Command @InvokeParams -EA SilentlyContinue
            }
            Catch {
                $M = "Error in computer compliance: $_"
                Write-EventLog @EventErrorParams -Message $M
                Write-Warning $M
            }

            Write-EventLog @EventVerboseParams -Message "Script 'Test computer' done"
        }
        #endregion

        $CompliantComputersUseWMI, $CompliantComputersNonWMI = $CompliantComputers.Where( { $_.UseWMI }, 'Split')

        #region 'Remove printers'
        $Jobs = @($RemoveWorksheet | Group-Object ServerName).ForEach( {
                $InvokeParams = @{
                    ArgumentList = , $_.Group
                    ComputerName = $_.Name
                    AsJob        = $true
                }
                if ($CompliantComputersNonWMI.ComputerName -contains $_.Name) {
                    Write-EventLog @EventOutParams -Message "Start 'Remove printers' script on '$($_.Name)' for printers:`n`n$($_.Group.PrinterName -join "`n")"

                    $InvokeParams.FilePath = $ScriptRemovePrintersItem
                    $InvokeParams.JobName = 'RemovePrinters'

                    Invoke-Command @InvokeParams

                }
                elseif ($CompliantComputersUseWMI.ComputerName -contains $_.Name) {
                    Write-EventLog @EventOutParams -Message "Start 'Remove printers WMI' script on '$($_.Name)' for printers:`n`n$($_.Group.PrinterName -join "`n")"

                    $InvokeParams.FilePath = $ScriptRemovePrintersWMIItem
                    $InvokeParams.JobName = 'RemovePrintersWMI'

                    Invoke-Command @InvokeParams
                }
                else {
                    Write-Verbose "Computer '$($_.Name)' is not compliant"
                }
            })

        $RemovePrintersJobResult = @($Jobs | Wait-Job -EA Ignore | Receive-Job)
        #endregion

        #region Call the 'Add printers' script
        $Jobs = @($PrintersWorksheet | Group-Object ServerName).ForEach( {
                if ($CompliantComputersNonWMI.ComputerName -contains $_.Name) {
                    Write-EventLog @EventOutParams -Message "Start 'Add printers' script on '$($_.Name)' for printers:`n`n$($_.Group.PrinterName -join "`n")"

                    $InvokeParams = @{
                        FilePath     = $ScriptAddPrintersItem
                        ArgumentList = , $_.Group
                        ComputerName = $_.Name
                        JobName      = 'AddPrinters'
                        AsJob        = $true
                    }
                    Invoke-Command @InvokeParams
                }
                elseif ($CompliantComputersUseWMI.ComputerName -contains $_.Name) {
                    Write-Error "Adding printers is only supported on Window Server 2012 and higher, not on '$($_.Name)'"
                }
            })

        $AddPrintersJobResult = @($Jobs | Wait-Job -EA Ignore | Receive-Job)

        Write-EventLog @EventVerboseParams -Message "Script 'Add printers' done"
        #endregion

        #region Remove unused printer ports
        if ($CompliantComputersNonWMI) {
            Write-EventLog @EventOutParams -Message "Start 'Remove unused printer ports' script on:`n`n$($CompliantComputersNonWMI.ComputerName -join "`n")"

            Try {
                $InvokeParams = @{
                    FilePath     = $ScriptRemoveUnusedPortsItem
                    ComputerName = $CompliantComputersNonWMI.ComputerName
                    JobName      = 'RemoveUnusedPorts'
                    AsJob        = $true
                }
                $Jobs = Invoke-Command @InvokeParams
            }
            Catch {
                $M = "Error in removing unused printer ports: $_"
                Write-Warning $M
                Write-EventLog @EventErrorParams -Message $M
            }

            @($Jobs | Wait-Job -EA Ignore | Receive-Job)

            Write-EventLog @EventVerboseParams -Message "Script 'Remove unused printer ports' done"
        }
        #endregion

        Get-Job | Remove-Job -Force
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {
        $Intro = $AddPrintersSummaryTable = $RemovePrintersSummaryTable = $ExcelWorksheetDescription = $Subject = $null

        $MailParams = @{
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
            To        = $MailTo
            Bcc       = $ScriptAdmin
        }

        #region Export to Excel
        $ExcelParams = @{
            Path         = $LogFile + '.xlsx'
            AutoSize     = $true
            FreezeTopRow = $true
        }

        if (Test-Path -Path $ExcelParams.Path) {
            Write-Warning "Excel file '$($ExcelParams.Path)' exists already and will be removed"
            Remove-Item -LiteralPath $ExcelParams.Path -Force
        }

        if ($AddPrintersJobResult) {
            Write-EventLog @EventOutParams -Message "Export worksheet 'Printers'"

            $AddPrintersJobResult | Select-Object *, @{N = 'Action'; E = { $_.Action -join ', ' } }  -ExcludeProperty Action, PSShowComputerName, RunspaceId |
            Export-Excel @ExcelParams -WorksheetName 'Printers' -TableName  'Printers'

            $MailParams.Attachments = $ExcelParams.Path
        }

        if ($RemovePrintersJobResult) {
            Write-EventLog @EventOutParams -Message "Export worksheet 'Remove'"

            $RemovePrintersJobResult | Select-Object *, @{N = 'Action'; E = { $_.Action -join ', ' } },
            @{N = 'Error'; E = { $_.Action -join ', ' } }  -ExcludeProperty Action, Error, PSShowComputerName, RunspaceId |
            Export-Excel @ExcelParams -WorksheetName 'Remove' -TableName 'Remove'

            $MailParams.Attachments = $ExcelParams.Path
        }

        if ($Error) {
            Write-EventLog @EventOutParams -Message "Export worksheet 'Error'"

            $Error.Exception.Message.ForEach( {
                    Write-EventLog @EventErrorParams -Message $_
                })

            $Error.Exception.Message | Select-Object @{N = 'Error message'; E = { $_ } } |
            Export-Excel @ExcelParams -WorksheetName 'Errors' -TableName 'Errors'

            $MailParams.Attachments = $ExcelParams.Path
        }
        #endregion

        #region Add summary table Printers
        $AddPrintersJobResultCountTotal = $AddPrintersJobResult.Count
        $AddPrintersJobResultCountError = @($AddPrintersJobResult.Where( { $_.Status -eq 'Error' })).Count
        $AddPrintersJobResultCountOk = @($AddPrintersJobResult.Where( { $_.Status -eq 'Ok' })).Count
        $AddPrintersJobResultCountInstalled = @($AddPrintersJobResult.Where( { $_.Status -eq 'Installed' })).Count
        $AddPrintersJobResultCountUpdated = @($AddPrintersJobResult.Where( { $_.Status -eq 'Updated' })).Count

        if ($AddPrintersJobResult) {
            $AddPrintersSummaryTable = "
            <p><i>Worksheet 'Printers':</i></p>
            <table>
                <tr><th>Quantity</th><th>Status</th></tr>
                $(if ($AddPrintersJobResultCountOk) {"<tr><td style=``"text-align: center``">$AddPrintersJobResultCountOk</td><td>OK</td></tr>"})
                $(if ($AddPrintersJobResultCountInstalled) {"<tr><td style=``"text-align: center``">$AddPrintersJobResultCountInstalled</td><td>Installed</td></tr>"})
                $(if ($AddPrintersJobResultCountUpdated) {"<tr><td style=``"text-align: center``">$AddPrintersJobResultCountUpdated</td><td>Updated</td></tr>"})
                $(if ($AddPrintersJobResultCountError) {"<tr><td style=``"text-align: center``">$AddPrintersJobResultCountError</td><td>Error</td></tr>"})
                <tr><td style=`"text-align: center`"><b>$AddPrintersJobResultCountTotal</b></td><b>Total</b></tr>
            </table>"
        }
        #endregion

        #region Add summary table Remove
        $RemovePrintersCountTotal = $RemovePrintersJobResult.Count
        $RemovePrintersCountError = @($RemovePrintersJobResult.Where( { $_.Status -eq 'Error' })).Count
        $RemovePrintersCountRemoved = @($RemovePrintersJobResult.Where( { $_.Status -eq 'Removed' })).Count

        if ($RemovePrintersJobResult) {
            $RemovePrintersSummaryTable = "
            <p><i>Worksheet 'Remove':</i></p>
            <table>
                <tr><th>Quantity</th><th>Status</th></tr>
                $(if ($RemovePrintersCountRemoved) {"<tr><td style=``"text-align: center``">$RemovePrintersCountRemoved</td><td>Removed</td></tr>"})
                $(if ($RemovePrintersCountError) {"<tr><td style=``"text-align: center``">$RemovePrintersCountError</td><td>Error</td></tr>"})
                <tr><td style=`"text-align: center`"><b>$RemovePrintersCountTotal</b></td><b>Total</b></tr>
            </table>"
        }
        #endregion

        #region Format mail message, subject and priority
        if ($Error) {
            $Subject = "FAILURE - $AddPrintersJobResultCountTotal printer queues, $($Error.Count) {0}" -f $(if ($Error.Count -eq 1) { 'error' }else { 'errors' })

            $Intro = "Failed to add/update printer queues due to <b>$($Error.Count) {0}</b> that {1} encountered during execution." -f $(if ($Error.Count -eq 1) { 'error' }else { 'errors' }), $(if ($Error.Count -eq 1) { 'was' }else { 'were' })

            $ExcelWorksheetDescription = "<p><i>* Please verify the worksheet 'Errors' in attachment</i></p>"
        }
        elseif ($AddPrintersJobResultCountError -or $RemovePrintersCountError) {
            $TotalErrorCount = $AddPrintersJobResultCountError + $RemovePrintersCountError
            $Subject = "FAILURE - $AddPrintersJobResultCountTotal printer queues, {0} {1}" -f
            $($TotalErrorCount), $(if ($Error.Count -eq 1) { 'error' }else { 'errors' })

            if ($AddPrintersJobResultCountError -and $RemovePrintersCountError) {
                $Intro = "Failed to add/update <b>$AddPrintersJobResultCountError printer queues</b> amd failed to remove <b>$RemovePrintersCountError printer queues</b>."
                $ExcelWorksheetDescription = "<p><i>* Please verify the 'Error' column in the worksheets 'Printers' and the worksheet 'Remove' in attachment.</i></p>"
            }
            elseif ($AddPrintersJobResultCountError) {
                $Intro = "Failed to add/update <b>$AddPrintersJobResultCountError printer queues</b>."
                $ExcelWorksheetDescription = "<p><i>* Please verify the 'Error' column in the worksheet 'Printers' in attachment.</i></p>"
            }
            elseif ($RemovePrintersCountError) {
                $Intro = "Failed to remove <b>$RemovePrintersCountError printer queues</b>."
                $ExcelWorksheetDescription = "<p><i>* Please verify the 'Error' column in the worksheet 'Remove' in attachment.</i></p>"
            }

        }
        elseif ($AddPrintersJobResultCountTotal -eq $AddPrintersJobResultCountOk) {
            $Subject = "$AddPrintersJobResultCountTotal printer queues, all correct"

            $Intro = "All printer queues are correct, no changes done."

            $ExcelWorksheetDescription = "<p><i>* Please find the overview in attachment.</i></p>"
        }
        else {
            $Subject = "$AddPrintersJobResultCountTotal printer queues"

            if ($AddPrintersJobResultCountInstalled -and $AddPrintersJobResultCountUpdated) {
                $Intro = "Successfully <b>installed $AddPrintersJobResultCountInstalled</b> and <b>updated $AddPrintersJobResultCountUpdated</b> printer queues."
                $Subject = "$AddPrintersJobResultCountTotal printer queues, $AddPrintersJobResultCountInstalled installed, $AddPrintersJobResultCountUpdated updated"
            }
            if ($AddPrintersJobResultCountInstalled -and (-not $AddPrintersJobResultCountUpdated)) {
                $Intro = "Successfully <b>installed $AddPrintersJobResultCountInstalled</b> {0}" -f
                $(if ($AddPrintersJobResultCountInstalled -eq 1) { 'printer queue.' }else { 'printer queues.' })

                $Subject = "$AddPrintersJobResultCountTotal printer queues, $AddPrintersJobResultCountInstalled added"
            }
            if ((-not $AddPrintersJobResultCountInstalled) -and $AddPrintersJobResultCountUpdated) {
                $Intro = "Successfully <b>updated $AddPrintersJobResultCountUpdated</b> {0}" -f
                $(if ($AddPrintersJobResultCountUpdated -eq 1) { 'printer queue.' }else { 'printer queues.' })

                $Subject = "$AddPrintersJobResultCountTotal printer queues, $AddPrintersJobResultCountUpdated updated"
            }

            $ExcelWorksheetDescription = "<p><i>* Please find in attachment an overview. All changes can be found in the field 'Action' of the worksheet 'Printers'."
        }

        $MailParams.Subject = $Subject
        $MailParams.Priority = if ($Error -or $AddPrintersJobResultCountError -or $RemovePrintersCountError) { 'High' } else { 'Normal' }
        $MailParams.Message = "
                $Intro
                $AddPrintersSummaryTable
                $RemovePrintersSummaryTable
                $ExcelWorksheetDescription
        "
        #endregion

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}
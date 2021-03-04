#Requires -Modules Pester, PrintManagement
#Requires -Version 5.1

BeforeAll {
    $testPrinter = [PSCustomObject]@{
        Status          = $null
        PrinterName     = 'PesterTestPrinter'
        ServerName      = 'Server1'
        DriverName      = 'KONICA MINOLTA Universal PCL v3.4'
        PortName        = 'PesterTestPort'
        PortHostAddress = '10.10.10.10'
        Collate         = $true
        Color           = $false
        DuplexingMode   = 'TwoSidedLongEdge'
        PaperSize       = 'A4'
        ShareName       = 'PesterTestPrinter'
        Shared          = $true
        Published       = $true
        KeepPrintedJobs = $false
        Location        = $null
        Comment         = 'My remarks'
        Action          = @()
        Error           = $null
    }
    
    # $testDriver = 'HP Universal Printing PCL 6'
    $testDriver ='HP Officejet Pro L7700 Series'
   
    Function Remove-TestPrintersHC {
        Get-Printer -Name "$($testPrinter.PrinterName)*" -EA Ignore | Remove-Printer -EA Ignore
    }
    Function Remove-TestPrinterPortsHC {
        $RetryTimes = 17
        $R = 0
        $Done = $false
    
        while (($R -le $RetryTimes) -and (-not $Done)) {
            Try {
                Get-PrinterPort -Name "$($testPrinter.PortName)*" -EA Ignore | Remove-PrinterPort -EA Stop
                $Done = $true
            }
            Catch {
                $R += 1
                Start-Sleep -Seconds 1
                Write-Verbose "Retry port removal $R"
            }
        }
    
        if (-not $Done) {
            throw $Error[0]
        }
    }
    
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    Remove-TestPrintersHC
    Remove-TestPrinterPortsHC
}
Describe 'Pester pre tests' {
    Context 'Drivers' {
        It 'the test driver should be different' {
            $testPrinter.DriverName | Should -Not -Be $testDriver
        }
        It 'test drivers should exist on the system' {
            $testDrivers = @(Get-PrinterDriver)

            $testDrivers.Name | Should -Contain $testDriver
            $testDrivers.Name | Should -Contain $testPrinter.DriverName
        }
    }
}
Describe 'Drivers' {
    Context 'when the driver is not installed' {
        BeforeAll {
            Remove-TestPrintersHC
            Remove-PrinterDriver -Name $testDriver -EA Ignore
            Get-PrinterDriver -Name $testDriver -EA Ignore | 
            Should -BeNullOrEmpty

            $testP1 = Copy-ObjectHC -Name $testPrinter
            $testP1.DriverName = $testDriver

            Mock Add-Printer
            Mock Add-PrinterDriver -ParameterFilter { $Name -eq $testDriver }
            Mock Get-PrintConfiguration {
                [PSCustomObject]@{
                    PaperSize = 'A4'
                    Color     = $true
                }
            }
            Mock Get-Printer {
                [PSCustomObject]@{
                    Comment  = $null
                    Location = $null
                }
            }
            Mock Set-PrintConfiguration
            Mock Set-Printer

            $testParams = @{
                Printers = $testP1
            }
            .$testScript @testParams
        }
        It 'install the driver from the store' {
            Should -Invoke Add-PrinterDriver -Times 1 -Exactly -Scope Context
        }
        It "add action 'Installed driver'" {
            $Printers.Action | Should -Contain 'Installed driver'
        }
        It 'no error is added' {
            $Printers.Error | Where-Object { $_ -like '*driver*' } | 
            Should -BeNullOrEmpty
        }
    }
    Context 'when the driver is already installed' {
        BeforeAll {
            Add-PrinterDriver -Name $testPrinter.DriverName -EA Ignore
            Get-PrinterDriver -Name $testPrinter.DriverName -EA Ignore | Should -Not -BeNullOrEmpty

            Mock Add-Printer
            Mock Add-PrinterDriver -ParameterFilter { $Name -eq $testPrinter.DriverName }
            Mock Get-PrintConfiguration {
                [PSCustomObject]@{
                    PaperSize = 'A4'
                    Color     = $true
                }
            }
            Mock Get-Printer {
                [PSCustomObject]@{
                    Comment  = $null
                    Location = $null
                }
            }
            Mock Set-PrintConfiguration
            Mock Set-Printer

            $testParams = @{
                Printers = Copy-ObjectHC -Name $testPrinter
            }
            .$testScript @testParams
        }
        It "it's not reinstalled" {
            Should -Not -Invoke Add-PrinterDriver -Scope Context
        }
        It 'no action is added' {
            $Printers.Action | Should -Not -Contain 'Installed driver'
        }
        It 'no error is added' {
            $Printers.Error | Where-Object { $_ -like '*driver*' } | 
            Should -BeNullOrEmpty
        }
    }
    Context 'when the driver is not found in the driver store' {
        BeforeAll {
            Mock Add-Printer
            Mock Add-PrinterDriver { throw 'Failed' }
            Mock Get-PrintConfiguration {
                [PSCustomObject]@{
                    PaperSize = 'A4'
                    Color     = $true
                }
            }
            Mock Get-Printer {
                [PSCustomObject]@{
                    Comment  = $null
                    Location = $null
                }
            }
            Mock Set-PrintConfiguration
            Mock Set-Printer

            $testP1 = Copy-ObjectHC -Name $testPrinter
            $testP1.DriverName = 'UnknownPrinterDriver'

            $testP2 = Copy-ObjectHC -Name $testPrinter
            $testP2.PrinterName += '2'
            $testP2.DriverName = 'UnknownPrinterDriver'

            $testP3 = Copy-ObjectHC -Name $testPrinter
            $testP3.PrinterName += '3'
            $testP3.DriverName = 'UnknownPrinterDriver2'

            $testParams = @{
                Printers = $testP1, $testP2, $testP3
            }
            .$testScript @testParams
        }
        It 'try to install the driver' {
            Should -Invoke Add-PrinterDriver -Times 2 -Exactly -Scope Context
        }
        It 'no action is added when it fails' {
            $Printers.Action | Should -Not -Contain 'Installed driver'
        }
        It 'an error is added for each printer' {
            $Printers.Error | Where-Object { $_ -like '*driver*' } | 
            Should -HaveCount 3
        }
    }
}
Describe 'Printer ports' {
    Context 'a correct printer port is' {
        BeforeAll {
            Mock Add-Printer
            Mock Get-PrintConfiguration {
                [PSCustomObject]@{
                    PaperSize = 'A4'
                    Color     = $true
                }
            }
            Mock Get-Printer {
                [PSCustomObject]@{
                    Comment  = $null
                    Location = $null
                }
            }
            Mock Set-PrintConfiguration
            Mock Set-Printer

            Remove-TestPrintersHC
            Remove-PrinterPort -Name $testPrinter.PortName -EA Ignore
            Add-PrinterPort -Name $testPrinter.PortName -PrinterHostAddress $testPrinter.PortHostAddress

            Mock Add-PrinterPort
            Mock Remove-PrinterPort

            $testParams = @{
                Printers = Copy-ObjectHC -Name $testPrinter
            }
            .$testScript @testParams
        }
        It 'not reinstalled' {
            Should -Not -Invoke Add-PrinterPort -Scope Context
        }
        It 'not removed' {
            Should -Not -Invoke Remove-PrinterPort -Scope Context
        }
        It 'no action is added' {
            $Printers.Action | Where-Object { $_ -like '*printer port*' } | Should -BeNullOrEmpty
        }
    }
    Context 'a missing printer port is' {
        BeforeAll {
            Mock Add-Printer
            Mock Get-PrintConfiguration {
                [PSCustomObject]@{
                    PaperSize = 'A4'
                    Color     = $true
                }
            }
            Mock Get-Printer {
                [PSCustomObject]@{
                    Comment  = $null
                    Location = $null
                }
            }
            Mock Set-PrintConfiguration
            Mock Set-Printer

            $testP1 = Copy-ObjectHC -Name $testPrinter
            $testP2 = Copy-ObjectHC -Name $testPrinter
            $testP2.PrinterName += '2'

            Remove-TestPrintersHC
            Remove-PrinterPort -Name $testPrinter.PortName -EA Ignore
            Get-PrinterPort -Name $testPrinter.PortName -EA Ignore | 
            Should -BeNullOrEmpty

            $testParams = @{
                Printers = $testP1, $testP2
            }
            .$testScript @testParams
        }

        It 'installed with the correct PostHostAddress' {
            $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
            $testResult.Name | Should -Be $testPrinter.PortName
            $testResult.PrinterHostAddress | 
            Should -Be $testPrinter.PortHostAddress
        }
        It 'registered as an action' {
            $Printers | 
            Where-Object { $_.Action -contains 'Installed printer port' } | 
            Should -Not -BeNullOrEmpty
        }
    }
    Context 'incorrect printer ports are' {
        Context 'when they are not in use' {
            BeforeAll {
                Mock Add-Printer
                Mock Get-PrintConfiguration {
                    [PSCustomObject]@{
                        PaperSize = 'A4'
                        Color     = $true
                    }
                }
                Mock Get-Printer {
                    [PSCustomObject]@{
                        Comment  = $null
                        Location = $null
                    }
                }
                Mock Set-PrintConfiguration
                Mock Set-Printer

                Remove-TestPrintersHC
                Remove-TestPrinterPortsHC

                $IncorrectPortHostAddress = $testPrinter.PortHostAddress + '1'

                Add-PrinterPort -Name $testPrinter.PortName -PrinterHostAddress $IncorrectPortHostAddress
                $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
                $testResult.Name | Should -Be $testPrinter.PortName
                $testResult.PrinterHostAddress | Should -Be $IncorrectPortHostAddress

                $testParams = @{
                    Printers = Copy-ObjectHC -Name $testPrinter
                }
                .$testScript @testParams
            }

            It 'removed and recreated' {
                $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
                $testResult.Name | Should -Be $testPrinter.PortName
                $testResult.PrinterHostAddress | 
                Should -Be $testPrinter.PortHostAddress
            }
            It 'registered as an action when recreated' {
                $Printers | Where-Object { $_.Action -contains 'Corrected printer port' } |
                Should -Not -BeNullOrEmpty
            }
        }
        Context 'when they are in use' {
            Context 'only by printers in the import file' {
                BeforeAll {
                    Mock Add-Printer
                    Mock Get-PrintConfiguration {
                        [PSCustomObject]@{
                            PaperSize = 'A4'
                            Color     = $true
                        }
                    }
                    Mock Get-Printer {
                        [PSCustomObject]@{
                            Comment  = $null
                            Location = $null
                        }
                    }
                    Mock Set-PrintConfiguration
                    Mock Set-Printer

                    $testP1 = Copy-ObjectHC -Name $testPrinter
                    $testP2 = Copy-ObjectHC -Name $testPrinter
                    $testP2.PrinterName += '2'

                    #region Create port with the same name but an inoccrect IP
                    Remove-TestPrintersHC
                    Remove-TestPrinterPortsHC

                    $IncorrectPortHostAddress = $testPrinter.PortHostAddress + '1'

                    Add-PrinterPort -Name $testPrinter.PortName -PrinterHostAddress $IncorrectPortHostAddress

                    $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
                    $testResult.Name | Should -Be $testPrinter.PortName
                    $testResult.PrinterHostAddress | 
                    Should -Be $IncorrectPortHostAddress
                    #endregion

                    #region Create 2 printers from the import file
                    $testPrintParams = @{
                        Name       = $testPrinter.PrinterName
                        DriverName = $testPrinter.DriverName
                        PortName   = $testPrinter.PortName
                    }
                    Remove-Printer -Name $testPrintParams.Name -EA Ignore
                    Add-Printer @testPrintParams

                    $testPrintParams.Name += '2'
                    Remove-Printer -Name $testPrintParams.Name -EA Ignore
                    Add-Printer @testPrintParams
                    #endregion

                    $testParams = @{
                        Printers = $testP1, $testP2
                    }
                    .$testScript @testParams
                }

                It 'corrected by being removed and recreated' {
                    $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
                    $testResult.Name | Should -Be $testPrinter.PortName
                    $testResult.PrinterHostAddress | 
                    Should -Be $testPrinter.PortHostAddress
                }
                It 'registered as an action when they are corrected' {
                    $Printers.Action | Should -Contain 'Corrected printer port'
                }
            }
            Context 'by installed printers that are not in the import file' {
                BeforeAll {
                    #region Create port with the same name but an inoccrect IP
                    Remove-TestPrintersHC
                    Remove-TestPrinterPortsHC

                    $IncorrectPortHostAddress = $testPrinter.PortHostAddress + '1'

                    Add-PrinterPort -Name $testPrinter.PortName -PrinterHostAddress $IncorrectPortHostAddress
                    $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
                    $testResult.Name | Should -Be $testPrinter.PortName
                    $testResult.PrinterHostAddress | 
                    Should -Be $IncorrectPortHostAddress
                    #endregion

                    #region Create 2 printers from the import file
                    $testPrintParams = @{
                        Name       = $testPrinter.PrinterName
                        DriverName = $testPrinter.DriverName
                        PortName   = $testPrinter.PortName
                    }
                    Add-Printer @testPrintParams

                    $testPrintParams.Name += '2'
                    Add-Printer @testPrintParams
                    #endregion

                    #region Create 1 printer not in the import file
                    $testPrintParams.Name += '3'
                    Remove-Printer -Name $testPrintParams.Name -EA Ignore
                    Add-Printer @testPrintParams
                    #endregion

                    Mock Add-Printer
                    Mock Add-PrinterPort
                    Mock Get-PrintConfiguration {
                        [PSCustomObject]@{
                            PaperSize = 'A4'
                            Color     = $true
                        }
                    }

                    Mock Set-PrintConfiguration
                    Mock Remove-Printer
                    Mock Remove-PrinterPort

                    $testP1 = Copy-ObjectHC -Name $testPrinter
                    $testP2 = Copy-ObjectHC -Name $testPrinter
                    $testP2.PrinterName += '2'

                    $testParams = @{
                        Printers = $testP1, $testP2
                    }
                    .$testScript @testParams
                }

                It 'not corrected or removed' {
                    @('Add-PrinterPort', 'Remove-PrinterPort', 'Remove-Printer').ForEach( {
                            Should -Not -Invoke -CommandName $_ -Scope Context
                        })
                }
                It 'registered as an error when not corrected' {
                    $Printers | Where-Object Error -Like 'Port in use by*' |
                    Should -HaveCount 2
                }
                It 'the installed port is left untouchted' {
                    $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
                    $testResult.Name | Should -Be $testPrinter.PortName
                    $testResult.PrinterHostAddress | Should -Be $IncorrectPortHostAddress
                }
                It 'the installed printer is left untouchted' {
                    $testResult = Get-Printer -Name $testPrintParams.Name -EA Ignore
                    $testResult.Name | Should -Be $testPrintParams.Name
                    $testResult.PortName | Should -Be $testPrinter.PortName
                }
            }
        }
    }
}
Describe 'Printers' {
    Context 'a missing printer is' {
        BeforeAll {
            Remove-TestPrintersHC
            Remove-TestPrinterPortsHC

            $testParams = @{
                Printers = Copy-ObjectHC -Name $testPrinter
            }
            .$testScript @testParams
        }

        It 'installed' {
            $testResult = Get-Printer -Name $testPrinter.PrinterName -EA Ignore
            $testResult.Name | Should -Be $testPrinter.PrinterName
            $testResult.DriverName | Should -Be $testPrinter.DriverName
            $testResult.PortName | Should -Be $testPrinter.PortName
        }
        It "add action 'Added printer'" {
            $Printers.Action | Should -Contain 'Added printer'
        }
        It 'configured with Set-PrintConfiguration' {
            $testResult = Get-PrintConfiguration -PrinterName $testPrinter.PrinterName -EA Ignore
            $testResult.Collate | Should -Be $testPrinter.Collate
            $testResult.Color | Should -Be $testPrinter.Color
            $testResult.DuplexingMode | Should -Be $testPrinter.DuplexingMode
            $testResult.PaperSize | Should -Be $testPrinter.PaperSize
        }
        It 'configured with Set-Printer' {
            $testResult = Get-Printer -Name $testPrinter.PrinterName -EA Ignore
            $testResult.Location | Should -Be $testPrinter.Location
            $testResult.KeepPrintedJobs | Should -Be $testPrinter.KeepPrintedJobs
            $testResult.ShareName | Should -Be $testPrinter.ShareName
            $testResult.Shared | Should -Be $testPrinter.Shared
            $testResult.Comment | Should -Be $testPrinter.Comment
        }
    }
    Context 'a correctly installed and configured printer is' {
        BeforeAll {
            Mock Add-Printer
            Mock Add-PrinterPort
            Mock Set-PrintConfiguration
            Mock Set-Printer
            Mock Remove-Printer
            Mock Remove-PrinterPort

            $testParams = @{
                Printers = Copy-ObjectHC -Name $testPrinter
            }
            .$testScript @testParams
        }

        It 'not reinstalled or reconfigured' {
            @(
                'Add-Printer', 'Add-PrinterPort',
                'Remove-Printer', 'Remove-PrinterPort',
                'Set-Printer', 'Set-PrintConfiguration'
            ).ForEach( {
                    Should -Invoke -CommandName $_ -Times 0 -Exactly -Scope Context
                })
        }
        It 'no action is added' {
            $Printers.Action | Should -BeNullOrEmpty
        }
        It 'no error is added' {
            $Printers.Error | Should -BeNullOrEmpty
        }
        It 'still installed' {
            $testResult = Get-Printer -Name $testPrinter.PrinterName -EA Ignore
            $testResult.Name | Should -Be $testPrinter.PrinterName
            $testResult.DriverName | Should -Be $testPrinter.DriverName
            $testResult.PortName | Should -Be $testPrinter.PortName
        }
    }
    Context 'an incorrectly configured printer with' {
        Context 'an incorrect port name is' {
            BeforeAll {
                #region Create port with an incorrect PortHostAddress
                Remove-TestPrintersHC
                Remove-TestPrinterPortsHC

                $IncorrectPortHostAddress = $testPrinter.PortHostAddress + '1'

                Add-PrinterPort -Name $testPrinter.PortName -PrinterHostAddress $IncorrectPortHostAddress
                $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
                $testResult.Name | Should -Be $testPrinter.PortName
                $testResult.PrinterHostAddress | 
                Should -Be $IncorrectPortHostAddress
                #endregion

                #region Create printer from the import file with an incorrect port
                $testPrintParams = @{
                    Name       = $testPrinter.PrinterName
                    DriverName = $testPrinter.DriverName
                    PortName   = $testPrinter.PortName
                }
                Remove-Printer -Name $testPrintParams.Name -EA Ignore
                Add-Printer @testPrintParams
                #endregion

                Mock Set-PrintConfiguration

                $testParams = @{
                    Printers = Copy-ObjectHC -Name $testPrinter
                }
                .$testScript @testParams
            }

            It 'removed and recreated with the correct port name' {
                $testResult = Get-Printer -Name $testPrinter.PrinterName -EA Ignore
                $testResult.Name | Should -Be $testPrinter.PrinterName
                $testResult.PortName | Should -Be $testPrinter.PortName

                $testResult = Get-PrinterPort -Name $testPrinter.PortName -EA Ignore
                $testResult.Name | Should -Be $testPrinter.PortName
                $testResult.PrinterHostAddress | 
                Should -Be $testPrinter.PortHostAddress
            }
            It 'registered as an action when corrected' {
                $Printers.Action | Should -Contain 'Added printer'
                $Printers.Action | Should -Contain 'Corrected printer port'
            }
            It 'configured with the correct settings' {
                Should -Invoke -CommandName 'Set-PrintConfiguration' -Times 1 -Exactly -Scope Context
            }
        }
        Context 'an incorrect driver is' {
            BeforeAll {
                #region Add wrong driver
                Add-PrinterDriver -Name $testDriver -EA Ignore
                Set-Printer -Name $testPrinter.PrinterName -DriverName $testDriver

                $testResult = Get-Printer -Name $testPrinter.PrinterName -EA Ignore
                $testResult.DriverName | Should -Not -Be $testPrinter.DriverName
                #endregion

                $testParams = @{
                    Printers = Copy-ObjectHC -Name $testPrinter
                }
                .$testScript @testParams
            }
            It 'corrected' {
                $testResult = Get-Printer -Name $testPrinter.PrinterName -EA Ignore
                $testResult.DriverName | Should -Be $testPrinter.DriverName
            }
            It 'and registered as an action' {
                $Printers.Action | Should -Contain 'Corrected driver'
            }
        }
        Context 'incorrect configuration is corrected and registered' {
            BeforeAll {
                #region Add incorrect config
                $testSetConfig = @{
                    PrinterName   = $testPrinter.PrinterName
                    PaperSize     = 'A2'
                    Color         = $true
                    Collate       = $true
                    DuplexingMode = 'OneSided'
                }
                Set-PrintConfiguration @testSetConfig

                $testSetConfig = @{
                    Name            = $testPrinter.PrinterName
                    Shared          = $false
                    ShareName       = $null
                    Published       = $false
                    KeepPrintedJobs = $true
                    Location        = 'A'
                }
                Set-Printer @testSetConfig
                #endregion

                $testP1 = Copy-ObjectHC -Name $testPrinter
                $testP1.PaperSize = 'A4'
                $testP1.Color = $false
                $testP1.Collate = $false
                $testP1.DuplexingMode = 'TwoSidedLongEdge'
                $testP1.Shared = $true
                $testP1.ShareName = 'PesterPrinter'
                $testP1.Location = 'B'
                $testP1.Published = $true
                $testP1.Comment = 'Remarks go here'

                $testParams = @{
                    Printers = $testP1
                }
                .$testScript @testParams

                $testPrintConfig = Get-PrintConfiguration -PrinterName $testPrinter.PrinterName
                $testGetPrinter = Get-Printer -Name $testPrinter.PrinterName
            }
            It 'Color' {
                $testPrintConfig.Color | Should -Be $testP1.Color
                $Printers.Action | Should -Contain 'Set color'
            }
            It 'Collate' {
                $testPrintConfig.Collate | Should -Be $testP1.Collate
                $Printers.Action | Should -Contain 'Set collate'
            }
            It 'DuplexingMode' {
                $testPrintConfig.DuplexingMode | Should -Be $testP1.DuplexingMode
                $Printers.Action | Should -Contain 'Set duplexingmode'
            }
            It 'PaperSize' {
                $testPrintConfig.PaperSize | Should -Be $testP1.PaperSize
                $Printers.Action | Should -Contain 'Set papersize'
            }
            It 'Shared' {
                $testGetPrinter.Shared | Should -Be $testP1.Shared
                $Printers.Action | Should -Contain 'Set shared'
            }
            It 'ShareName' {
                $testGetPrinter.ShareName | Should -Be $testP1.ShareName
                $Printers.Action | Should -Contain 'Set sharename'
            }
            It 'Location' {
                $testGetPrinter.Location | Should -Be $testP1.Location
                $Printers.Action | Should -Contain 'Set location'
            }
            It 'Published' {
                $testGetPrinter.Published | Should -Be $testP1.Published
                $Printers.Action | Should -Contain 'Set published'
            }
            It 'Comment' {
                $testGetPrinter.Comment | Should -Be $testP1.Comment
                $Printers.Action | Should -Contain 'Set Comment'
            }
        }
    }
}
Describe 'Status' {
    Context "is set to 'Installed' when" {
        It 'the printer was not on the system before but is after the run' {
            Remove-TestPrintersHC

            $testParams = @{
                Printers = Copy-ObjectHC -Name $testPrinter
            }
            .$testScript @testParams

            $Printers.Status | Should -Be 'Installed'
        }
    }
    Context "is set to 'Ok' when" {
        It 'the printer was on the system and is correctly configured' {
            $testParams = @{
                Printers = Copy-ObjectHC -Name $testPrinter
            }
            .$testScript @testParams

            $Printers.Status | Should -Be 'Ok'
        }
    }
    Context "is set to 'Updated' when" {
        It 'the printer had an incorrect configuration' {
            Set-PrintConfiguration -PrinterName $testPrinter.PrinterName -PaperSize A2

            $testParams = @{
                Printers = Copy-ObjectHC -Name $testPrinter
            }
            .$testScript @testParams

            $Printers.Status | Should -Be 'Updated'
        }
    }
    Context "is set to 'Error' when" {
        It 'the printer had cannot be updated correctly' {
            $testNewPrinter = Copy-ObjectHC -Name $testPrinter
            $testNewPrinter.DriverName = 'NonExisting'

            $testParams = @{
                Printers = $testNewPrinter
            }
            .$testScript @testParams

            $Printers.Status | Should -Be 'Error'
        }
    }
}

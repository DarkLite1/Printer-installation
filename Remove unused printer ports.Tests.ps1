#Requires -Modules Pester, PrintManagement
#Requires -Version 5.1

BeforeAll {
    $testDriverName = @(Get-PrinterDriver)[0].Name

    $testInstalledPrinters = @(
        [PSCustomObject]@{
            PrinterName     = 'PesterTestPrinter1'
            DriverName      = $testDriverName
            PortName        = 'PesterTestPort1'
            PortHostAddress = '192.168.1.1'
        }
        [PSCustomObject]@{
            PrinterName     = 'PesterTestPrinter2'
            DriverName      = $testDriverName
            PortName        = 'PesterTestPort2'
            PortHostAddress = '192.168.1.2'
        }
        [PSCustomObject]@{
            PrinterName     = 'PesterTestPrinter3'
            DriverName      = $testDriverName
            PortName        = 'PesterTestPort2'
            PortHostAddress = '192.168.1.2'
        }
    )
    
    $testPortsNotInUse = @(
        [PSCustomObject]@{
            PortName           = 'PesterTestPort3'
            PrinterHostAddress = '192.168.1.3'
        }
        [PSCustomObject]@{
            PortName           = 'PesterTestPort4'
            PrinterHostAddress = '192.168.1.4'
        }
        [PSCustomObject]@{
            PortName           = 'PesterTestPort5'
            PrinterHostAddress = 'printer5'
        }
        [PSCustomObject]@{
            PortName           = 'PesterTestPort6:'
            PrinterHostAddress = $null
        }
    )

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
}

Describe 'when a printer port' {
    BeforeAll {
        #region Remove printer queues
        Get-Printer | Where-Object {
            ($_.Name -like 'Pester*') -and
            (
                ($testInstalledPrinters.PortName -contains $_.PortName) -or
                ($testPortsNotInUse.PortName -contains $_.PortName)
            )
        } | ForEach-Object {
            Remove-Printer -Name $_.Name
        }
        #endregion

        #region Remove and add printer ports
        @($testInstalledPrinters + $testPortsNotInUse) | Select-Object PortName -Unique |
        ForEach-Object {
            Remove-PrinterPort -Name $_.PortName -EA Ignore

            $AddPortParams = @{
                Name = $_.PortName
            }

            if ($_.PrinterHostAddress) {
                $AddPortParams.PrinterHostAddress = $_.PrinterHostAddress
            }

            Add-PrinterPort @AddPortParams
        }
        #endregion

        #region Add printer queues
        $testInstalledPrinters.ForEach( {
                $testPrintParams = @{
                    Name       = $_.PrinterName
                    DriverName = $_.DriverName
                    PortName   = $_.PortName
                }
                Add-Printer @testPrintParams
            })
        #endregion
    }
    It 'is not in use it is removed' {
        . $testScript

        Get-PrinterPort | 
        Where-Object { $testPortsNotInUse.PortName -Contains $_.Name } |
        Should -BeNullOrEmpty
    }
    It 'is in use it is not removed' {
        . $testScript

        Get-PrinterPort | 
        Where-Object { $testInstalledPrinters.PortName -Contains $_.Name } |
        Should -Not -BeNullOrEmpty
    }
}
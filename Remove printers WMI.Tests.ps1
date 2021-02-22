#Requires -Modules Pester, PrintManagement
#Requires -Version 5.1

BeforeAll {
    $testDriverName = @(Get-PrinterDriver)[0].Name

    $testPrinters = @(
        [PSCustomObject]@{
            PrinterName = 'PesterTestPrinter1'
            Status      = $null
            Action      = @()
            Error       = $null
        }
        [PSCustomObject]@{
            PrinterName = 'PesterTestPrinter2a'
            Status      = $null
            Action      = @()
            Error       = $null
        }
        [PSCustomObject]@{
            PrinterName = 'PesterTestPrinter2b'
            Status      = $null
            Action      = @()
            Error       = $null
        }
    )
    
    $testInstalledPrinters = @(
        [PSCustomObject]@{
            PrinterName     = $testPrinters[0].PrinterName
            DriverName      = $testDriverName
            PortName        = 'PesterTestPort1'
            PortHostAddress = '5.10.10.1'
        }
        [PSCustomObject]@{
            PrinterName     = $testPrinters[1].PrinterName
            DriverName      = $testDriverName
            PortName        = 'PesterTestPort2'
            PortHostAddress = '5.10.10.2'
        }
        [PSCustomObject]@{
            PrinterName     = $testPrinters[2].PrinterName
            DriverName      = $testDriverName
            PortName        = 'PesterTestPort2'
            PortHostAddress = '5.10.10.2'
        }
    )

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    Mock Invoke-Sqlcmd2
    Mock Send-MailHC
    Mock Write-EventLog
}

Describe 'when a printer' {
    BeforeAll {
        #region Install all printers
        $testInstalledPrinters.ForEach( {
                Remove-Printer -Name $_.PrinterName -EA Ignore
            })
        $testInstalledPrinters | Select-Object PortName -Unique | ForEach-Object {
            Remove-PrinterPort -Name $_.PortName -EA Ignore
        }
        $testInstalledPrinters | Sort-Object PortName -Unique | ForEach-Object {
            Add-PrinterPort -Name $_.PortName -PrinterHostAddress $_.PortHostAddress
        }
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
    Context "has a print port that is not in use by another printer" {
        BeforeAll {
            $testParams = @{
                Printers = Copy-ObjectHC $testPrinters[0]
            }
            .$testScript @testParams
        }
        It 'the printer queue is removed' {
            Get-Printer | 
            Where-Object Name -EQ $testInstalledPrinters[0].PrinterName | 
            Should -BeNullOrEmpty
        }
        It 'the printer port is removed' {
            Get-PrinterPort | 
            Where-Object Name -EQ $testInstalledPrinters[0].PortName | 
            Should -BeNullOrEmpty
        }
        It "'Action' contains 'Removed printer queue'" {
            $Printers.Action | Should -Contain 'Removed printer queue'
        }
        It "'Action' contains 'Removed printer port'" {
            $Printers.Action | 
            Should -Contain "Removed port '$($testInstalledPrinters[0].PortName) > $($testInstalledPrinters[0].PortHostAddress)'"
        }
        It "'Status' is 'Removed'" {
            $Printers.Status | Should -Be 'Removed'
        }
        It "'Error' is empty" {
            $Printers.Error | Should -BeNullOrEmpty
        }
    }
    Context "has a print port that is in use by another printer" {
        BeforeAll {
            $testParams = @{
                Printers = Copy-ObjectHC $testPrinters[1]
            }
            .$testScript @testParams
        }
        It 'the printer queue is removed' {
            Get-Printer | Where-Object Name -EQ $testInstalledPrinters[1].PrinterName | Should -BeNullOrEmpty
        }
        It 'the printer port is not removed' {
            Get-PrinterPort | 
            Where-Object Name -EQ $testInstalledPrinters[1].PortName | 
            Should -Not -BeNullOrEmpty
        }
        It "'Error' is 'Port in use'" {
            $Printers.Error | 
            Should -Contain "Port '$($testInstalledPrinters[1].PortName) > $($testInstalledPrinters[1].PortHostAddress)' cannot be removed because it is in use by printers '$($testInstalledPrinters[2].PrinterName)'"
        }
        It "'Action' is not containing port removal" {
            $Printers.Action | Where-Object { $_ -like '*port*' } | 
            Should -BeNullOrEmpty
        }
        It "'Action' contains 'Removed printer queue'" {
            $Printers.Action | Should -Contain 'Removed printer queue'
        }
        It "'Status' is 'Error' because the port couldn't be removed" {
            $Printers.Status | Should -Be 'Error'
        }
    }
    Context 'is not found on the print server' {
        BeforeAll {
            Mock Remove-Printer
            Mock Remove-PrinterPort

            $testP1 = Copy-ObjectHC $testPrinters[1]
            $testP1.PrinterName = 'NotExistingPrinterName'

            $testParams = @{
                Printers = $testP1
            }
            .$testScript @testParams
        }
        It "no queue is removed" {
            Should -Not -Invoke Remove-Printer -Scope Context
        }
        It "no port is removed" {
            Should -Not -Invoke Remove-PrinterPort -Scope Context
        }
        It "'Error' is 'Printer queue not found'" {
            $Printers.Error | Should -Contain 'Printer queue not found'
        }
        It "'Action' is empty because nothing happened" {
            $Printers.Action | Should -BeNullOrEmpty
        }
        It "'Status' is 'Error' because the print queue wasn't found" {
            $Printers.Status | Should -Be 'Error'
        }
    }
}

#Requires -Modules Pester, PrintManagement
#Requires -Version 5.1

BeforeAll {
    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and ($Subject -eq 'FAILURE')
    }
    
    $MailUsersParams = {
        ($To -eq $MailTo) -and ($Priority -eq 'High') -and ($Subject -like 'FAILURE - Incorrect input')
    }
    
    $testPrintersWorksheet = @(
        [PSCustomObject]@{
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
            Published       = $false
            KeepPrintedJobs = $false
            Location        = $null
            Comment         = 'My remarks'
        }
    )
    
    $testRemoveWorksheet = @(
        [PSCustomObject]@{
            ServerName  = 'Server1'
            PrinterName = 'Printer1'
        }
        [PSCustomObject]@{
            ServerName  = 'Server2'
            PrinterName = 'Printer2'
        }
    )

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        MailTo     = 'BobLeeSwagger@shooter.net'
        ImportFile = New-Item "TestDrive:/Printers.xlsx" -ItemType File
        LogFolder  = New-Item "TestDrive:/log" -ItemType Directory
    }

    Mock Import-Excel
    Mock Invoke-Command
    Mock Send-MailHC
    Mock Write-EventLog
}

Describe 'Import file' {
    BeforeAll {
        Mock Export-Excel
    }
    Context 'send no error mail when' {
        It 'a non mandatory property is missing' {
            Mock Import-Excel {
                $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                $testP1.PSObject.Properties.Remove('Color')
                $testP1.PSObject.Properties.Remove('Shared')
                $testP1
            } -ParameterFilter {
                $WorksheetName -eq 'Printers'
            }

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*boolean property*")
            }
            Should -Not -Invoke Send-MailHC -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*missing property*")
            }
        } 
        Context 'DuplexingMode is correct' {
            BeforeAll {
                $forEach = @(
                    $null,
                    'OneSided',
                    'TwoSidedShortEdge',
                    'TwoSidedLongEdge'
                ).ForEach( { @{Name = $_ } })
            }

            It '<Name>' -forEach $forEach {
                Param (
                    $Name
                )

                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.DuplexingMode = $Name
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Not -Invoke Send-MailHC -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*DuplexingMode*")
                }
            }
        }
        Context 'PaperSize is correct' {
            BeforeAll {
                $forEach = @(
                    $null,
                    'a4',
                    'letter',
                    'A3'
                ).ForEach( { @{Name = $_ } })
            }

            It '<Name>' -forEach $forEach {
                Param (
                    $Name
                )

                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.PaperSize = $Name
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Not -Invoke Send-MailHC -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*PaperSize*")
                }
            }
        }
    }
    Context 'send an error mail to the admin when' {
        BeforeEach {
            Mock Import-Excel  -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }
            Mock Import-Excel  -ParameterFilter {
                $WorksheetName -eq 'Printers'
            }
        }
        It 'the file is not found' {
            $testNewParams = $testParams.Clone()
            $testNewParams.ImportFile = 'NotFound.xlsx'
            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Import file*not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It "the worksheet 'Printers' is missing" {
            Mock Import-Excel {
                throw 'Worksheet not found'
            } -ParameterFilter {
                $WorksheetName -eq 'Printers'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like '*Worksheet not found*')
            }
        }
        It "the worksheet 'Remove' is missing" {
            Mock Import-Excel {
                throw 'Worksheet not found'
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like '*Worksheet not found*')
            }
        } 
    }
}
Describe 'Printers worksheet' {
    BeforeAll {
        Mock Import-Excel
    }
    Context 'string properties' {
        It 'leading and trailing spaces are removed' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                $testF1.Comment = ' TEST  '
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Printers'
            }

            .$testScript @testParams

            $PrintersWorksheet.Comment | Should -Be  'TEST'
        }
    }
    Context 'send an error mail to the users when' {
        It 'Shared is used instead of ShareName only' {
            Mock Import-Excel {
                $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                $testP1 | 
                Add-Member -NotePropertyName Shared -NotePropertyValue $true
                $testP1
            } -ParameterFilter {
                $WorksheetName -eq 'Printers'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like "*Shared*is not valid*")
            }
        }
        Context 'a mandatory property is missing' {
            It '<name>' -forEach @(
                @{name = 'DriverName'},
                @{name = 'PortHostAddress'},
                @{name = 'PortName'},
                @{name = 'PrinterName'},
                @{name = 'ServerName'}
            ) {
                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.PSObject.Properties.Remove($name)
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.$name = $null
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 2 -ParameterFilter {
                    (&$MailUsersParams) -and 
                    ($Message -like "*missing property '$name'*")
                }
                Should -Invoke Write-EventLog -Exactly 2 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            } -Tag test
        } 
        Context 'a boolean property is not TRUE, FALSE or NULL' {
            It '<_>' -forEach @(
                @{name = 'Color'},
                @{name = 'Collate'},
                @{name = 'KeepPrintedJobs'},
                @{name = 'Published'}
            ) {
                foreach ($testVal in @($true, $false, $null)) {
                    Mock Import-Excel {
                        $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                        $testP1.$name = $testVal
                        $testP1
                    } -ParameterFilter {
                        $WorksheetName -eq 'Printers'
                    }

                    .$testScript @testParams

                    Should -Not -Invoke Send-MailHC -ParameterFilter {

                        (&$MailAdminParams) -and 
                        ($Message -like "*boolean property*")
                    }
                }

                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.$name = 'Wrong'
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and 
                    ($Message -like "*boolean property*")
                }
            }
        }
        Context 'conflicting Excel rows are found' {
            It 'duplicate ServerName PrinterName combination' {
                Mock Import-Excel {
                    Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    Copy-ObjectHC -Name $testPrintersWorksheet[0]
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*duplicate PrinterName*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'different PortHostAddress for the same PortName on the same server' {
                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.PortHostAddress = 'A'
                    $testP1

                    $testP2 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP2.PrinterName += '2'
                    $testP2.PortHostAddress = 'B'
                    $testP2
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*conflicting PortHostAddress*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
        Context 'a mandatory property is missing' {
            It '<Name>' -forEach @(
                @{name = 'ServerName'},
                @{name = 'PrinterName'},
                @{name = 'DriverName'}
            ) {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testF1.$name = $null
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and 
                    ($Message -like "*is missing property '$name'*")
                }
            } 
        }
        Context 'there is a space in' {
            It '<Name>' -forEach @(
                @{name = 'ServerName'},
                @{name = 'PrinterName'},
                @{name = 'ShareName'}
            ) {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testF1.$Name = "Test wrong"
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*$Name*cannot contain spaces*")
                }
            } 
        }
        Context 'incorrect value for' {
            It 'DuplexingMode' {
                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.DuplexingMode = 'Wrong'
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*DuplexingMode*")
                }
            }
            It 'PaperSize' {
                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.PaperSize = 'Wrong'
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*PaperSize*")
                }
            }
        }
    }
    Context 'set the property value' {
        Context 'Shared' {
            It 'to true when a ShareName is given' {
                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.ShareName = 'PRINTER'
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                $PrintersWorksheet.Shared | Should -BeTrue
            }
            It 'to false when there is no ShareName' {
                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                    $testP1.ShareName = $null
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Printers'
                }

                .$testScript @testParams

                $PrintersWorksheet.Shared | Should -BeFalse
            }
        }
    }
}
Describe 'Remove worksheet' {
    BeforeAll {
        Mock Import-Excel
    }
    Context 'string properties' {
        It 'leading and trailing spaces are removed' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.PrinterName = ' TEST  '
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            $RemoveWorksheet.PrinterName | Should -Be  'TEST'
        }
    }
    Context 'send an error mail to the users when' {
        It "a duplicate PrinterName ServerName combo is found" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.PrinterName = 'Printer'
                $testF1.ServerName = 'Server'
                $testF1

                $testF2 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF2.PrinterName = 'Printer'
                $testF2.ServerName = 'Server'
                $testF2
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*Duplicate PrinterName, ServerName*')
            }
        }
        It "the worksheet Printers contains the same PrinterName/ServerName combo" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.PrinterName = 'Printer'
                $testF1.ServerName = 'Server'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
                $testF1.PrinterName = 'Printer'
                $testF1.ServerName = 'Server'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Printers'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like "*Printers*Remove*contain duplicate records*")
            }
        }
        Context 'a mandatory property is missing' {
            BeforeAll {
                $forEach = @(
                    'PrinterName',
                    'ServerName'
                ).ForEach( { @{Name = $_ } })
            }

            It '<Name>' -forEach $forEach {
                Param (
                    $Name
                )
                Mock Import-Excel {
                    $testP1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                    $testP1.$Name = $null
                    $testP1
                } -ParameterFilter {
                    $WorksheetName -eq 'Remove'
                }
                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*mandatory property '$Name' is missing*")
                }
            }
        }
        Context 'there is a space in' {
            BeforeAll {
                $forEach = @(
                    'ServerName',
                    'PrinterName'
                ).ForEach( { @{Name = $_ } })
            }
            It '<Name>' -forEach $forEach {
                Param (
                    $Name
                )

                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                    $testF1.$Name = "Test wrong"
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'Remove'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*$Name*cannot contain spaces*")
                }
            } 
        }
    }
}
Describe 'Test computer script' {
    BeforeAll {
        Mock Export-Excel
    }
    It 'send an error mail to the admin when the file is not found' {
        $testNewParams = $testParams.Clone()
        $testNewParams.ScriptTestComputer = 'NotFound.ps1'
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*script file*not found*")
        }
    }
    It 'is called for all ServerNames in Add/Remove worksheets' {
        Mock Import-Excel {
            $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
            $testP1.ServerName = 'S1'
            $testP1

            $testP2 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
            $testP2.ServerName = 'S2'
            $testP2

            $testP3 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
            $testP3.ServerName = 'S3'
            $testP3

            $testP4 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
            $testP4.ServerName = 'S3'
            $testP4.PrinterName += '1'
            $testP4
        } -ParameterFilter {
            $WorksheetName -eq 'Printers'
        }

        Mock Import-Excel {
            $testP1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP1.ServerName = 'S4'
            $testP1

            $testP2 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP2.ServerName = 'S4'
            $testP2.PrinterName += '1'
            $testP2

            $testP3 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP3.ServerName = 'S5'
            $testP3
        } -ParameterFilter {
            $WorksheetName -eq 'Remove'
        }

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
            $FilePath -eq $ScriptTestComputerItem

        }
    }
}
Describe 'Add printers script' {
    BeforeAll {
        Mock Export-Excel
    }
    It 'send an error mail to the admin when the file is not found' {
        $testNewParams = $testParams.Clone()
        $testNewParams.ScriptAddPrinters = 'NotFound.ps1'
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*script file*not found*")
        }
    }
    It 'is called once for each ServerName' {
        Mock Import-Excel {
            $testP1 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
            $testP1.ServerName = 'S1'
            $testP1

            $testP2 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
            $testP2.ServerName = 'S2'
            $testP2

            $testP3 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
            $testP3.ServerName = 'S3'
            $testP3

            $testP4 = Copy-ObjectHC -Name $testPrintersWorksheet[0]
            $testP4.ServerName = 'S3'
            $testP4.PrinterName += '1'
            $testP4
        } -ParameterFilter {
            $WorksheetName -eq 'Printers'
        }
        Mock Invoke-Command {
            @(
                [PSCustomObject]@{
                    ComputerName = 'S1'
                    UseWMI       = $false
                }
                [PSCustomObject]@{
                    ComputerName = 'S2'
                    UseWMI       = $false
                }
                [PSCustomObject]@{
                    ComputerName = 'S3'
                    UseWMI       = $false
                }
            )
        } -ParameterFilter {
            $FilePath -eq $ScriptTestComputerItem
        }

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 3 -Exactly -ParameterFilter {
            $JobName -eq 'AddPrinters'
        }
    }
}
Describe 'Remove printers script' {
    BeforeAll {
        Mock Export-Excel
    }
    It 'send an error mail to the admin when the file is not found' {
        $testNewParams = $testParams.Clone()
        $testNewParams.ScriptRemovePrinters = 'NotFound.ps1'
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*script file*not found*")
        }
    }
    It 'is called once for each ServerName' {
        Mock Import-Excel {
            $testP1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP1.ServerName = 'S1'
            $testP1

            $testP2 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP2.ServerName = 'S2'
            $testP2
        } -ParameterFilter {
            $WorksheetName -eq 'Remove'
        }
        Mock Invoke-Command {
            @(
                [PSCustomObject]@{
                    ComputerName = 'S1'
                    UseWMI       = $false
                }
                [PSCustomObject]@{
                    ComputerName = 'S2'
                    UseWMI       = $false
                }
            )
        } -ParameterFilter {
            $FilePath -eq $ScriptTestComputerItem
        }

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 2 -Exactly -ParameterFilter {
            $JobName -eq 'RemovePrinters'
        }
    }
}
Describe 'Remove printers WMI script' {
    BeforeAll {
        Mock Export-Excel
    }
    It 'send an error mail to the admin when the file is not found' {
        $testNewParams = $testParams.Clone()
        $testNewParams.ScriptRemovePrintersWMI = 'NotFound.ps1'
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*script file*not found*")
        }
    }
    It 'is called once for each ServerName' {
        Mock Import-Excel {
            $testP1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP1.ServerName = 'S1'
            $testP1

            $testP2 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP2.ServerName = 'S2'
            $testP2

            $testP2 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP2.ServerName = 'S3'
            $testP2
        } -ParameterFilter {
            $WorksheetName -eq 'Remove'
        }
        Mock Invoke-Command {
            @(
                [PSCustomObject]@{
                    ComputerName = 'S1'
                    UseWMI       = $true
                }
                [PSCustomObject]@{
                    ComputerName = 'S2'
                    UseWMI       = $true
                }
                [PSCustomObject]@{
                    ComputerName = 'S3'
                    UseWMI       = $false
                }
            )
        } -ParameterFilter {
            $FilePath -eq $ScriptTestComputerItem
        }

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 2 -Exactly -ParameterFilter {
            $JobName -eq 'RemovePrintersWMI'
        }
    }
}
Describe 'Remove unused printer ports script' {
    BeforeAll {
        Mock Export-Excel
    }
    It 'send an error mail to the admin when the file is not found' {
        $testNewParams = $testParams.Clone()
        $testNewParams.ScriptRemoveUnusedPorts = 'NotFound.ps1'
        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*script file*not found*")
        }
    }
    It 'is called once for each ServerName' {
        Mock Import-Excel {
            $testP1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP1.ServerName = 'S1'
            $testP1

            $testP2 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testP2.ServerName = 'S2'
            $testP2
        } -ParameterFilter {
            $WorksheetName -eq 'Remove'
        }
        Mock Invoke-Command {
            @(
                [PSCustomObject]@{
                    ComputerName = 'S1'
                    UseWMI       = $false
                }
                [PSCustomObject]@{
                    ComputerName = 'S2'
                    UseWMI       = $false
                }
            )
        } -ParameterFilter {
            $FilePath -eq $ScriptTestComputerItem
        }

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
            $JobName -eq 'RemoveUnusedPorts'
        }
    }
}
Describe 'create an Excel file in the log folder with' {
    BeforeAll {
        Mock Export-Excel
    }
    It "the worksheet 'Printers'" {
        Mock Invoke-Command {
            @(
                [PSCustomObject]@{
                    ComputerName = $testPrintersWorksheet[0].ServerName
                    UseWMI       = $false
                }
            )
        } -ParameterFilter {
            $FilePath -eq $ScriptTestComputerItem
        }
        Mock Invoke-Command {
            Start-Job -ScriptBlock {
                Copy-ObjectHC -Name $Using:testPrintersWorksheet[0]
            }
        } -ParameterFilter {
            $JobName -eq 'AddPrinters'
        }
        Mock Import-Excel {
            Copy-ObjectHC -Name $testPrintersWorksheet[0]
        } -ParameterFilter {
            $WorksheetName -eq 'Printers'
        }

        .$testScript @testParams

        Should -Invoke Export-Excel -Times 1 -Exactly -ParameterFilter {
            $WorksheetName -eq 'Printers'
        }
    }
    It "the worksheet 'Errors' when there are errors found" {
        Mock Invoke-Command {
            throw 'Planned test error'
        } -ParameterFilter {
            $FilePath -eq $ScriptTestComputerItem
        }
        Mock Import-Excel {
            Copy-ObjectHC -Name $testPrintersWorksheet[0]
        } -ParameterFilter {
            $WorksheetName -eq 'Printers'
        }

        .$testScript @testParams

        Should -Invoke Export-Excel -Times 1 -Exactly -ParameterFilter {
            $WorksheetName -eq 'Errors'
        }
    }
}
Describe 'Send an e-mail to the user' {
    It 'with an Excel file in attachment' {
        Mock Invoke-Command {
            @(
                [PSCustomObject]@{
                    ComputerName = $testPrintersWorksheet[0].ServerName
                    UseWMI       = $false
                }
            )
        } -ParameterFilter {
            $FilePath -eq $ScriptTestComputerItem
        }
        Mock Invoke-Command {
            Start-Job -ScriptBlock {
                Copy-ObjectHC -Name $Using:testPrintersWorksheet[0]
            }
        } -ParameterFilter {
            $JobName -eq 'AddPrinters'
        }
        Mock Import-Excel {
            Copy-ObjectHC -Name $testPrintersWorksheet[0]
        } -ParameterFilter {
            $WorksheetName -eq 'Printers'
        }

        .$testScript @testParams

        Should -Invoke Send-MailHC -Times 1 -Exactly -ParameterFilter {
            $Attachments -ne $mull
        }
    }
}
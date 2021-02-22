#Requires -Modules Pester, PrintManagement
#Requires -Version 5.1

BeforeAll {
    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and ($Subject -eq 'FAILURE')
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    # Enable mocking of in script functions    
    Function Test-IsAdminHC { }
    Function Test-PowerShellVersionHC { }
}

Describe 'a terminating error is thrown when' {
    It 'when we are not local admin' {
        Mock Get-Module { $true }
        Mock Test-IsAdminHC { $false }
        Mock Test-PowerShellVersionHC { $true }

        ( { .$testScript } | Should -Throw -PassThru).Exception |
        Should -BeLike  "*Administrator*"
    }
    It 'PowerShell 5.1 or higher is not installed' {
        Mock Get-Module { $true }
        Mock Test-IsAdminHC { $true }
        Mock Test-PowerShellVersionHC { $false }

        ( { .$testScript } | Should -Throw -PassThru).Exception |
        Should -BeLike  "*PowerShell*"
    }
}
Describe 'when there is no terminating error the output has the property' {
    It 'ComputerName' {
        Mock Get-Module { $true }
        Mock Test-IsAdminHC { $true }
        Mock Test-PowerShellVersionHC { $true }

        (.$testScript).ComputerName | Should -Be $env:COMPUTERNAME
    }
    Context 'UseWMI' {
        It 'set to TRUE when the PrintManagement module is installed' {
            Mock Get-Module { $false }
            Mock Test-IsAdminHC { $true }
            Mock Test-PowerShellVersionHC { $true }

            (.$testScript).UseWMI | Should -BeTrue
        }
        It 'set to FALSE when it is not installed' {
            Mock Get-Module { $true }
            Mock Test-IsAdminHC { $true }
            Mock Test-PowerShellVersionHC { $true }

            (.$testScript).UseWMI | Should -BeFalse
        }
    }

}

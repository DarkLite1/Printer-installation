#Requires -Modules PrintManagement
#Requires -Version 5.1

<#
.SYNOPSIS
    Test the server to see if it's compliant.

.DESCRIPTION
    This script is intended to run on a local server. It will verify the server 
    to see if it meets the minimum requirements to run PowerShell code.

    the require statements at the top is not supported with 
    `Invoke-Command -FilePath`. That's why the 'Begin' clause repeat the tests.
    https://stackoverflow.com/questions/51185882/invoke-command-ignores-requires-in-the-script-file
#>

[CmdLetBinding()]
Param ()

Begin {
    Function Test-IsAdminHC {
        <#
        .SYNOPSIS
            Check if a user is local administrator.

        .DESCRIPTION
            Check if a user is member of the local group 'Administrators' and 
            returns TRUE if he is, FALSE if not.

        .EXAMPLE
            Test-IsAdminHC -SamAccountName SrvBatch
            Returns TRUE in case SrvBatch is admin on this machine

        .EXAMPLE
            Test-IsAdminHC
            Returns TRUE if the current user is admin on this machine
        #>

        [CmdLetBinding()]
        [OutputType([Boolean])]
        Param (
            $SamAccountName = [Security.Principal.WindowsIdentity]::GetCurrent()
        )

        Try {
            $Identity = [Security.Principal.WindowsIdentity]$SamAccountName
            $Principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $Identity
            $Result = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            Write-Verbose "Administrator permissions: $Result"
            $Result
        }
        Catch {
            throw "Failed to determine if the user '$SamAccountName' is local admin: $_"
        }
    }

    Function Test-PowerShellVersionHC {
        [CmdLetBinding()]
        [OutputType([Boolean])]
        Param (
            [Int]$Major,
            [Int]$Minor
        )

        (($Major -gt 5) -or (($Major -eq 5) -and ($Minor -ge 1)))
    }
}

Process {
    Try {
        $Obj = [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            UseWMI = $false
        }

        #region Test administrator privileges
        if (-not (Test-IsAdminHC)) {
            throw "Administrator privileges are required for user '$env:USERNAME@$env:USERDNSDOMAIN'."
        }
        #endregion

        #region Test PowerShell version
        if (-not (Test-PowerShellVersionHC -Major $PSVersionTable.PSVersion.Major -Minor $PSVersionTable.PSVersion.Minor)) {
            throw "PowerShell version 5.1 or higher is required."
        }
        #endregion

        #region Test module PrintManagement
        if (-not (Get-Module -ListAvailable -Name 'PrintManagement')) {
            $Obj.UseWMI = $true
        }
        #endregion

        $Obj
    }
    Catch {
        throw "Computer '$env:COMPUTERNAME' is not compliant: $_"
    }
}
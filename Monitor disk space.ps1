#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Send an e-mail with the free disk space on computers.

    .DESCRIPTION
        Send an e-mail with Excel file in attachment containing the drives found
        on specific computers and their free disk space.

    .PARAMETER ComputerName
        Collection of computer names to scan for hard drives.

    .PARAMETER ExcludeDrive
        Collection of drive letters to excluded from the report.

    .PARAMETER ColorFreeSpaceBelow
        Colors used in the Excel file for visually marking low disk space.
        Ex:
        - Red    : 10 > less than 10% free disk space is colored red
        - Orange : 15 > less than 15% free disk space is colored orange

    .PARAMETER SendMail.Header
        The header to use in the e-mail sent t the end user.

    .PARAMETER SendMail.To
        List of e-mail addresses where to send the e-mail too.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [string]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Monitor\Monitor disk space\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)
        
Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start
        $error.Clear()
        
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

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        
        $file = Get-Content -Path $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion
        
        #region Test .json file properties
        try {
            if (-not ($ComputerNames = $file.ComputerName)) {
                throw "Property 'ComputerName' not found."
            }
            if (-not ($ExcludedDrives = $file.ExcludeDrive)) {
                throw "Property 'ExcludeDrive' not found."
            }
            if (-not ($ColorFreeSpaceBelow = $file.ColorFreeSpaceBelow)) {
                throw "Property 'ColorFreeSpaceBelow' not found."
            }
            if (-not ($ColorFreeSpaceBelow -is [PSCustomObject])) {
                throw "Property 'ColorFreeSpaceBelow' is not a key value pair of a color with a percentage number."
            }
            $ColorFreeSpaceBelow.PSObject.Properties | ForEach-Object {
                if (-not ($_.Value -is [Int])) {
                    throw "Property 'ColorFreeSpaceBelow' with color '$($_.Name)' contains value '$($_.Value)' that is not a number."
                }
            }
            if (-not ($SendMail = $file.SendMail)) {
                throw "Property 'SendMail' not found."
            }
            if (-not $SendMail.To) {
                throw "Property 'SendMail.To' not found."
            }
            if (-not $SendMail.Header) {
                throw "Property 'SendMail.Header' not found."
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}       
Process {
    Try {
        
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
End {
    try {
        
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}
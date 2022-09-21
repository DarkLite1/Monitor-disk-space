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
        
        #region Create log folder
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
            $ComputerNames | Group-Object | Where-Object { $_.Count -ge 2 } | ForEach-Object {
                throw "Property 'ComputerName' contains the duplicate value '$($_.Name)'."
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
            $ExcludedDrives = foreach ($e in $file.ExcludeDrive) {
                if (-not $e.ComputerName) {
                    throw "A computer name is mandatory for an excluded drive. Use the wildcard '*' to excluded the drive letter for all computers."    
                }
                foreach ($d in $e.DriveLetter) {
                    if ($d -notMatch '^[A-Z]$' ) {
                        throw "Excluded drive letter '$d' is not a single alphabetical character"    
                    }
                    [PSCustomObject]@{
                        ComputerName = $e.ComputerName
                        DriveLetter  = '{0}:' -f $d.ToUpper()
                    }
                    
                    $M = "Exclude drive letter '$d' on computer '$($e.ComputerName)'"
                    Write-Verbose $M
                    Write-EventLog @EventVerboseParams -Message $M
                }
            }
            if ($ColorFreeSpaceBelow = $file.ColorFreeSpaceBelow) {
                if (-not ($ColorFreeSpaceBelow -is [PSCustomObject])) {
                    throw "Property 'ColorFreeSpaceBelow' is not a key value pair of a color with a percentage number."
                }
                $ColorFreeSpaceBelow.PSObject.Properties | ForEach-Object {
                    if (-not ($_.Value -is [Int])) {
                        throw "Property 'ColorFreeSpaceBelow' with color '$($_.Name)' contains value '$($_.Value)' that is not a number."
                    }
                }
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
        #region Get drives
        $M = 'Get hard disk details for {0} computers' -f $ComputerNames.Count
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $params = @{
            ClassName   = 'Win32_LogicalDisk'
            Filter      = 'DriveType = 3'
            ErrorAction = 'SilentlyContinue'
            Verbose     = $false
        }
        [array]$drives = foreach ($computer in $ComputerNames) {
            Write-Verbose "Get drives on computer '$computer'"
            Get-CimInstance @params -ComputerName $computer
        }
        #endregion

        #region Filter out excluded drives
        foreach ($e in $ExcludedDrives) {
            if ($e.ComputerName -eq '*') {
                $drives = $drives.Where({ $_.DeviceID -ne $e.DriveLetter })
            }
            else {
                $drives = $drives.Where({ 
                        -not (
                            ($_.PSComputerName -eq $e.ComputerName) -and
                            ($_.DeviceID -eq $e.DriveLetter)
                        )
                    }
                )
            }
        }
        #endregion

        $M = "Found '{0}' drives" -f $drives.Count
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
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
        $excelParams = @{
            Path         = "$LogFile.xlsx"
            AutoSize     = $true
            FreezeTopRow = $true
        }
        $mailParams = @{
            To        = $SendMail.To
            Bcc       = $ScriptAdmin
            Message   = $null
            Subject   = $null
            LogFolder = $LogParams.LogFolder
            Header    = $SendMail.Header
            Save      = "$LogFile - Mail.html"
        }

        #region Export data to Excel
        if ($drives) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Drives'

            $drives | Select-Object -Property @{
                Name       = 'ComputerName'
                Expression = { $_.PSComputerName }
            },
            @{
                Name       = 'Drive'
                Expression = { $_.DeviceID }
            },
            @{
                Name       = 'DriveName'
                Expression = { $_.VolumeName }
            },
            @{
                Name       = 'Size'
                Expression = { [Math]::Round( $_.Size / 1GB, 2) }
            },
            @{
                Name       = 'UsedSpace'
                Expression = { 
                    [Math]::Round(($_.Size - $_.FreeSpace) / 1GB, 2) 
                }
            },
            @{
                Name       = 'FreeSpace'
                Expression = { [Math]::Round( $_.FreeSpace / 1GB, 2) }
            },
            @{
                Name       = 'Free'
                Expression = { 
                    [Math]::Round( ($_.FreeSpace / $_.Size) * 100, 2) 
                }
            } |
            Export-Excel @excelParams -AutoNameRange -CellStyleSB {
                Param (
                    $WorkSheet,
                    $TotalRows,
                    $LastColumn
                )

                @(
                    $WorkSheet.Names[
                    'Size', 'FreeSpace', 'UsedSpace'
                    ].Style).ForEach( {
                        $_.NumberFormat.Format = '?\ \G\B'
                    }
                )

                $WorkSheet.Cells.Style.HorizontalAlignment = 'Center'
            }

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Count results, errors, ...
        $counter = @{
            drives    = ($drives | Measure-Object).Count
            computers = ($ComputerNames | Measure-Object).Count
            errors    = ($Error.Exception.Message | Measure-Object).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'

        $mailParams.Subject = '{0} computer{1}, {2} drive{3}' -f
        $counter.computers,
        $(if ($counter.computers -ne 1) { 's' }), 
        $counter.drives,
        $(if ($counter.drives -ne 1) { 's' })

        #endregion

        if ($counter.errors) {
            #region Export errors to Excel
            $excelParams.WorksheetName = $excelParams.TableName = 'Errors'

            $Error.Exception.Message | Select-Object -Unique | 
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
            #endregion
            
            #region Mail subject, priority, message
            $mailParams.Priority = 'High'

            $mailParams.Subject += ', {0} error{1}' -f $counter.errors, $(
                if ($counter.errors -ne 1) { 's' }
            )
            $mailParams.Message = "<p>Detected <b>{0} non terminating error{1}</b></p>" -f $counter.errors, 
            $(
                if ($counter.errors -gt 1) { 's' }
            )
            #endregion
        }

        #region Send mail
        $mailParams.Message += "
            <p>Scan results of the hard disks:</p>
            <table>
                <tr><th>Computers</th><td>{0}</td></tr>
                <tr><th>Drives</th><td>{1}</td></tr>
            </table>" -f 
        $counter.computers,
        $counter.drives

        if ($mailParams.Attachments) {
            $mailParams.Message += '<p><i>* Check the attachment for details</i></p>'
        }
        Get-ScriptRuntimeHC -Stop  
        Send-MailHC @mailParams
        #endregion
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
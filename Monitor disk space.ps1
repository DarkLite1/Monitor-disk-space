#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Scan computers for free disk space and create a report.

    .DESCRIPTION
        This script reads a .JSON input file containing all the required
        parameters (ComputerName, ...). Each computer is then scanned for its
        hard drives and an Excel file is created containing an overview of the
        drives found (drive letter, drive name, disk size, free space, ...).

        Check the Example.json file on how to create a correct input file. All
        available parameters in the input file are explained below.

    .PARAMETER ComputerName
        Collection of computer names to scan for hard drives.

    .PARAMETER ExcludeDrive
        Collection of drive letters to excluded from the report.

        "ExcludeDrive": [
            {
                "ComputerName": "*",
                "DriveLetter": ["S"]
            }
        ]
        For all computers (wildcard '*') exclude drive letter 'S'.

        "ExcludeDrive": [
            {
                "ComputerName": "PC1",
                "DriveLetter": ["B", "D"]
            }
        ]
        On computer 'PC1' exclude drive letters 'B' and 'D' .

    .PARAMETER ColorFreeSpaceBelow
        Defines the colors used in the Excel file to indicate low free disk space below a specific percentage or amount of GB.

        "ColorFreeSpaceBelow": {
            "Type": "GB",
            "Value": { "Red": 10, "Orange": 15 },
            "?": "Type: GB | %"
        },
        Color the rows with free space less than 15GB orange and 10GB red.

    .PARAMETER SendMail.Header
        The header to use in the e-mail sent to the users. If SendMail.Header
        is not provided the ScriptName will be used.

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
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start
        $error.Clear()

        #region Add color assembly
        Add-Type -Assembly System.Drawing
        Set-Culture 'en-US'
        #endregion

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
            if (-not ($SendMailHeader = $SendMail.Header)) {
                $SendMailHeader = $ScriptName
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

            $highlightExcelRow = [Ordered]@{ }

            if ($ColorFreeSpaceBelow = $file.ColorFreeSpaceBelow) {
                if (
                    (-not ($ColorFreeSpaceBelow -is [PSCustomObject])) -or
                    (-not $ColorFreeSpaceBelow.Type) -or
                    (-not $ColorFreeSpaceBelow.Value) -or
                    (-not ($ColorFreeSpaceBelow.Value -is [PSCustomObject]))
                ) {
                    throw "Property 'ColorFreeSpaceBelow' is not a valid object. A valid object has the format @{Type='GB'; Value=@{'Red'=10; 'Orange'=15}}."
                }

                if ($ColorFreeSpaceBelow.Type -notMatch '^GB$|^%$') {
                    throw "Property 'ColorFreeSpaceBelow' only supports type 'GB' or '%'."
                }

                foreach (
                    $property in
                    $ColorFreeSpaceBelow.Value.PSObject.Properties |
                    Sort-Object 'Value'
                ) {
                    try {
                        $null = $property.Value.ToInt16($null)
                    }
                    catch {
                        throw "Property 'ColorFreeSpaceBelow' with color '$($property.Name)' contains value '$($property.Value)' that is not a number."
                    }

                    Try {
                        $ColorValue = $property.Name
                        $null = [System.Drawing.Color]$property.Name
                    }
                    Catch {
                        Throw "Property 'ColorFreeSpaceBelow' with 'Color' value '$ColorValue' is not valid because it's not a proper color"
                    }

                    $highlightExcelRow.Add(
                        $property.Value, [System.Drawing.Color]$property.Name
                    )

                    $M = "Highlight Excel row with free space lower than '{0}{1}' in '{2}'" -f
                    $property.Value, $ColorFreeSpaceBelow.Type, $property.Name
                    Write-Verbose $M
                    Write-EventLog @EventVerboseParams -Message $M
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
            try {
                Write-Verbose "Get drives on computer '$computer'"
                Get-CimInstance @params -ComputerName $computer
            }
            catch {
                Write-Error "Failed getting drives on '$computer': $_"
                $Error.RemoveAt(1)
            }
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
            Header    = $SendMailHeader
            Save      = "$LogFile - Mail.html"
        }

        if ($drives) {
            #region Export data to Excel
            $excelParams.WorksheetName = $excelParams.TableName = 'Drives'

            $column = @{}

            if ($ColorFreeSpaceBelow.Type -eq 'GB') {
                $column.Color = 'F'
                $column.Sort = 'FreeSpace'
            }
            elseif ($ColorFreeSpaceBelow.Type -eq '%') {
                $column.Color = 'G'
                $column.Sort = 'Free'
            }
            else {
                $column.Sort = 'ComputerName'
            }

            [array]$drives = $drives |
            Select-Object -Property @{
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
            }

            $excelWorkbook = $drives | Sort-Object $column.Sort |
            Export-Excel @excelParams -PassThru -AutoNameRange -CellStyleSB {
                Param (
                    $workSheet,
                    $TotalRows,
                    $LastColumn
                )

                @(
                    $workSheet.Names[
                    'Size', 'FreeSpace', 'UsedSpace'
                    ].Style).ForEach( {
                        $_.NumberFormat.Format = '?\ \G\B'
                    }
                )

                @(
                    $workSheet.Names['Free'].Style).ForEach( {
                        $_.NumberFormat.Format = '? \%'
                    }
                )

                $workSheet.Cells.Style.HorizontalAlignment = 'Center'
            }
            #endregion

            $mailParams.Attachments = $excelParams.Path

            #region Format percentage and set row color
            if ($highlightExcelRow) {
                $workSheet = $excelWorkbook.Workbook.Worksheets[$excelParams.WorkSheetName]

                $conditionParams = @{
                    WorkSheet = $workSheet
                    Range     = '{0}2:{0}{1}' -f
                    $column.Color, $workSheet.Dimension.Rows
                }

                $firstTimeThrough = $true
                foreach ($h in $highlightExcelRow.GetEnumerator()) {
                    if ($firstTimeThrough) {
                        $firstTimeThrough = $False
                        Add-ConditionalFormatting @conditionParams -BackgroundColor $h.Value.Name -RuleType LessThan -ConditionValue $h.Name
                    }
                    else {
                        Add-ConditionalFormatting @conditionParams -BackgroundColor $h.Value.Name -RuleType Between -ConditionValue $h.Name -ConditionValue2 $previousValue
                    }

                    $previousValue = $h.Name
                }
            }
            #endregion

            $excelWorkbook.Save()
            $excelWorkbook.Dispose()
        }

        #region Count results, errors, ...
        $counter = @{
            computers = ($ComputerNames | Measure-Object).Count
            drives    = $drives.Count
            errors    = $Error.Count
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
            $excelParams.WorkSheetName ='Errors'
            $excelParams.TableName ='Errors'

            $Error.Exception.Message |
            Select-Object @{
                Name       = 'Error message'
                Expression = { $_ }
            } |
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
            #endregion

            #region Mail subject, priority, message
            $mailParams.Priority = 'High'

            $mailParams.Subject += ', {0} error{1}' -f $counter.errors, $(
                if ($counter.errors -ne 1) { 's' }
            )
            $mailParams.Message = "<p>Detected <b>{0} non terminating error{1}.</b></p>" -f $counter.errors,
            $(
                if ($counter.errors -gt 1) { 's' }
            )
            #endregion
        }

        #region Send mail
        $countedColorRows = if ($highlightExcelRow) {
            $previousValue = $null
            foreach ($h in $highlightExcelRow.GetEnumerator()) {
                '<tr><th>{0}</th><td>{1}</td></tr>' -f
                $(
                    if (-not $previousValue) {
                        'less than {0}{1}' -f $h.Key, $ColorFreeSpaceBelow.Type
                    }
                    else {
                        'between {0}{1} and {2}{1}' -f
                        $previousValue, $ColorFreeSpaceBelow.Type, $h.Key
                    }
                ),
                $(
                    $driveCounter = $drives.Where(
                        {
                            if (-not $previousValue) {
                                $_."$($column.Sort)" -lt $h.Key
                            }
                            else {
                                ($_."$($column.Sort)" -ge $previousValue) -and
                                ($_."$($column.Sort)" -lt $h.Key)
                            }
                        }
                    ).Count

                    if ($driveCounter -eq 1) {
                        '{0} drive' -f $driveCounter
                    }
                    else {
                        '{0} drives' -f $driveCounter
                    }
                )
                $previousValue = $h.Key
            }
        }

        $mailParams.Message += "
            <p>Scan results of the hard disks:</p>
            <table>
                <tr><th>Computers</th><td>{0}</td></tr>
                <tr><th>Drives</th><td>{1}</td></tr>
                $countedColorRows
            </table>" -f $counter.computers, $counter.drives


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
#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Get-CimInstance
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach 'ScriptName', 'ImportFile' {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE')
        }    
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx::\notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like "*Failed creating the log folder 'xxx::\notExistingLocation'*")
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'is missing property <_>' -ForEach @(
            'ComputerName', 'ExcludeDrive', 'ColorFreeSpaceBelow', 'SendMail'
        ) {
            $testJsonFile = @{
                ComputerName        = @('PC1', 'PC2')
                ExcludeDrive        = @('S')
                ColorFreeSpaceBelow = @{
                    Red    = 10
                    Orange = 15
                }
                SendMail            = @{
                    Header = 'Application X disc space report'
                    To     = 'bob@contoso.com'
                }
            }
            $testJsonFile.Remove($_)
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Property '$_' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'is missing property SendMail.To' {
            $testJsonFile = @{
                ComputerName        = @('PC1', 'PC2')
                ExcludeDrive        = @('S')
                ColorFreeSpaceBelow = @{
                    Red    = 10
                    Orange = 15
                }
                SendMail            = @{
                    Header = 'Application X disc space report'
                    # To     = 'bob@contoso.com'
                }
            }
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Property 'SendMail.To' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'is missing property SendMail.Header' {
            $testJsonFile = @{
                ComputerName        = @('PC1', 'PC2')
                ExcludeDrive        = @('S')
                ColorFreeSpaceBelow = @{
                    Red    = 10
                    Orange = 15
                }
                SendMail            = @{
                    # Header = 'Application X disc space report'
                    To = 'bob@contoso.com'
                }
            }
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Property 'SendMail.Header' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'the property ColorFreeSpaceBelow' {
            It 'is not a key value pair' {
                @{
                    ComputerName        = @("PC1", "PC2")
                    ExcludeDrive        = @("S")
                    ColorFreeSpaceBelow = 5
                    SendMail            = @{
                        Header = "Application X disc space report"
                        To     = "bob@contoso.com"
                    }
                } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

                .$testScript @testParams
                            
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*Property 'ColorFreeSpaceBelow' is not a key value pair of a color with a percentage number*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'is not a color with a number' {
                @{
                    ComputerName        = @("PC1", "PC2")
                    ExcludeDrive        = @("S")
                    ColorFreeSpaceBelow = @{
                        Red = 'text'
                    }
                    SendMail            = @{
                        Header = "Application X disc space report"
                        To     = "bob@contoso.com"
                    }
                } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

                .$testScript @testParams
                            
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*Property 'ColorFreeSpaceBelow' with color 'Red' contains value 'text' that is not a number*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
        It 'the property ComputerName contains duplicates' {
            $testJsonFile = @{
                ComputerName        = @('PC1', 'PC2', 'PC2')
                ExcludeDrive        = @('S')
                ColorFreeSpaceBelow = @{
                    Red    = 10
                    Orange = 15
                }
                SendMail            = @{
                    Header = 'Application X disc space report'
                    To     = 'bob@contoso.com'
                }
            }
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Property 'ComputerName' contains the duplicate value 'PC2'*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'the property ExcludeDrive contains an invalid string' {
            $testJsonFile = @{
                ComputerName        = @('PC1')
                ExcludeDrive        = @('dd')
                ColorFreeSpaceBelow = @{
                    Red    = 10
                    Orange = 15
                }
                SendMail            = @{
                    Header = 'Application X disc space report'
                    To     = 'bob@contoso.com'
                }
            }
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Excluded drive 'dd' is not a single alphabetical character*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        Mock Get-CimInstance {
            [PSCustomObject]@{
                PSComputerName = 'PC1'
                FreeSpace      = 1073741824
                Size           = 5368709120
                VolumeName     = 'OTHER'
                DeviceID       = 'A:'
            }
            [PSCustomObject]@{
                PSComputerName = 'PC1'
                FreeSpace      = 4523646976
                Size           = 5366607872
                VolumeName     = 'DATA'
                DeviceID       = 'B:'
            }
            [PSCustomObject]@{
                PSComputerName = 'PC1'
                FreeSpace      = 4523646976
                Size           = 5366607872
                VolumeName     = 'OS'
                DeviceID       = 'C:'
            }
        } -ParameterFilter {
            $ComputerName -eq 'PC1'
        }
        Mock Get-CimInstance {
            [PSCustomObject]@{
                PSComputerName = 'PC2'
                FreeSpace      = 53687091200
                Size           = 107374182400
                VolumeName     = 'BLA'
                DeviceID       = 'A:'
            }
        } -ParameterFilter {
            $ComputerName -eq 'PC2'
        }

        $testJsonFile = @{
            ComputerName        = @('PC1', 'PC2')
            ExcludeDrive        = @('B', 'c')
            ColorFreeSpaceBelow = @{
                Red    = 10
                Orange = 15
            }
            SendMail            = @{
                Header = 'Application X disc space report'
                To     = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    It 'call Get-CimInstance once for each computer' {
        @('PC1', 'PC2') | ForEach-Object {
            Should -Invoke Get-CimInstance -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ClassName -eq 'Win32_LogicalDisk') -and
                ($Filter -eq 'DriveType = 3') -and
                ($ComputerName -eq $_) -and
                ($ErrorAction -eq 'SilentlyContinue') 
            }
        }

        Should -Invoke Get-CimInstance -Times 2 -Exactly -Scope Describe
    }
    It 'ignore excluded drives' {
        $drives | Should -HaveCount 2
        @(
            [PSCustomObject]@{
                PSComputerName = 'PC1'
                DeviceID       = 'A:'
            }
            [PSCustomObject]@{
                PSComputerName = 'PC2'
                DeviceID       = 'A:'
            }
        ) | ForEach-Object {
            $drives.PSComputerName | Should -Contain $_.PSComputerName
            $drives.DeviceID | Should -Contain $_.DeviceID
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    ComputerName    = 'PC1'
                    Drive           = 'A:'
                    DriveName       = 'OTHER'
                    'Size(GB)'      = '5'
                    'UsedSpace(GB)' = '4'
                    'FreeSpace(GB)' = '1'
                    'FreeSpace(%)'  = '20'
                }
                @{
                    ComputerName    = 'PC2'
                    Drive           = 'A:'
                    DriveName       = 'BLA'
                    'Size(GB)'      = '100'
                    'UsedSpace(GB)' = '50'
                    'FreeSpace(GB)' = '50'
                    'FreeSpace(%)'  = '50'
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Drives'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.ComputerName -eq $testRow.ComputerName
                }
                $actualRow.Drive | Should -Be $testRow.Drive
                $actualRow.DriveName | Should -Be $testRow.DriveName
                $actualRow.'Size(GB)' | Should -Be $testRow.'Size(GB)'
                $actualRow.'FreeSpace(GB)' | Should -Be $testRow.'FreeSpace(GB)'
                $actualRow.'FreeSpace(%)' | Should -Be $testRow.'FreeSpace(%)'
                $actualRow.'UsedSpace(GB)' | Should -Be $testRow.'UsedSpace(GB)'
            }
        }
    }
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '1 file moved') -and
            ($Attachments -like '*log.xlsx') -and
            ($Message -like (
                "*From: <a href=`"{0}`">{0}</a><br>To: <a href=`"{1}`">{1}</a><br>Move files older than 3 days<br>Moved: 1*" -f $(
                    "\\$env:COMPUTERNAME\C$\$($testFolder.Source.Substring(3))"
                ),
                $(
                    "\\$env:COMPUTERNAME\C$\$($testFolder.Destination.Substring(3))"
                )
            ))
        }
    } -Skip
}
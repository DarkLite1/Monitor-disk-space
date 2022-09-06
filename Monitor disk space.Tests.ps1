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
    }
}
Describe 'when all tests pass' {
    It 'call Get-CimInstance once for each computer' {
        $testComputerNames = @('PC1', 'PC2')

        $testJsonFile = @{
            ComputerName        = $testComputerNames
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
                    
        $testComputerNames | ForEach-Object {
            Should -Invoke Get-CimInstance -Exactly 1 -ParameterFilter {
                ($ClassName -eq 'Win32_LogicalDisk') -and
                ($Filter -eq 'DriveType = 3') -and
                ($ComputerName -eq $_) -and
                ($ErrorAction -eq 'SilentlyContinue') 
            }
        }

        Should -Invoke Get-CimInstance -Exactly $testComputerNames.Count
    }
} -Tag test
#Requires -Modules Pester, PrintManagement
#Requires -Version 5.1

BeforeAll {
    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and ($Subject -eq 'FAILURE')
    }
    
    $testGetPrinter = @(
        [PSCustomObject]@{
            Name            = 'PesterTestPrinter1'
            Comment         = $null
            ComputerName    = 'PC1'
            DriverName      = 'KONICA MINOLTA 4700PSeries PCL6'
            ShareName       = $null
            PortName        = '192.168.1.1'
            Datatype        = 'RAW'
            KeepPrintedJobs = $false
            Published       = $false
            Priority        = 1
            Shared          = $false
            StartTime       = 0
            Type            = 'Local'
            UntilTime       = 0
            PrinterStatus   = 'Normal'
            JobCount        = 0
        }
    )
    
    $testGetPrinterDriver = @(
        [PSCustomObject]@{
            Name               = $testGetPrinter[0].DriverName
            ComputerName       = $testGetPrinter[0].ComputerName
            MajorVersion       = 4
            DriverVersion      = 1688859053146112
            Manufacturer       = 'Microsoft'
            PrinterEnvironment = 'Windows x64'
        }
    )
    
    $testGetPrintConfiguration = @(
        [PSCustomObject]@{
            PrinterName   = $testGetPrinter[0].Name
            ComputerName  = $testGetPrinter[0].ComputerName
            DuplexingMode = 'TwoSidedLongEdge'
            Color         = $false
            Collate       = $true
            PaperSize     = 'A4'
        }
    )
    
    $testGetPrinterPort = @(
        [PSCustomObject]@{
            Name               = $testGetPrinter[0].PortName
            ComputerName       = $testGetPrinter[0].ComputerName
            Description        = 'Standard TCP / IP Port'
            Protocol           = 'LPR'
            PrinterHostAddress = '192.168.1.1'
            PrinterHostIP      = $null
            PortNumber         = 515
            SNMPIndex          = 1
            SNMPCommunity      = 'public'
            SNMPEnabled        = 'False'
            LprQueueName       = 'LPR'
            LprByteCounting    = $false
        }
        [PSCustomObject]@{
            Name                = 'COM1:'
            ComputerName        = 'DEUSPFRAN0002'
            Description         = 'Local Port'
            PortMonitor         = 'Local Monitor'
            ElementName         = $null
            InstanceID          = $null
            CommunicationStatus = $null
            DetailedStatus      = $null
            HealthState         = $null
            InstallDate         = $null
            OperatingStatus     = $null
            OperationalStatus   = $null
            PrimaryStatus       = $null
            Status              = $null
            StatusDescriptions  = $null
        }
    )
    
    $testGetPrinterSNMPInfoHC = @(
        [PSCustomObject][Ordered]@{
            SNMP_ComputerName  = $testGetPrinterPort[0].PrinterHostAddress
            SNMP_Status        = 'Ok'
            SNMP_CommunityName = 'public'
            SNMP_Name          = $testGetPrinter.Name
            SNMP_Model         = 'KONICA MINOLTA 4700PSeries PCL6'
            SNMP_Contact       = $null
            SNMP_SN            = '1ead46840asdf'
            SNMP_Description   = 'Best printer in the world'
            SNMP_Location      = $null
            SNMP_UpTimeNIC     = $null
            SNMP_CountUnit     = $null
            SNMP_CountTotal    = $null
            SNMP_CountPowerOn  = $null
            SNMP_Toners        = $null
            SNMP_TonerColors   = $null
            SNMP_TonerNames    = $null
            SNMP_MaxSpeedUnit  = $null
            SNMP_MaxSpeed      = $null
            SNMP_SMTP          = $null
            SNMP_NICSpeedMbps  = $null
            SNMP_Alert         = $null
        }
    )
    
    $testGetDNSInfoHC = @(
        [PSCustomObject]@{
            ComputerName = $testGetPrinter[0].Name
            IP           = '192.168.2.2'
            HostName     = $testGetPrinter[0].Name
        }
        [PSCustomObject]@{
            ComputerName = $testGetPrinterPort[0].PrinterHostAddress
            IP           = '192.168.2.3'
            HostName     = $testGetPrinter[0].name
        }
    )

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        MailTo           = 'BobLeeSwagger@shooter.net'
        OU               = 'contoso.com'
        ScriptName       = 'Test'
        ComputersNotInOU = New-Item 'TestDrive:/ComputersNotInOU.txt' -ItemType File
        LogFolder        = New-Item 'TestDrive:/Log' -ItemType Directory
    }

    Mock Get-CimInstance {
        [PSCustomObject]@{
            Version = '6.3.1'
        }
    }
    Mock Get-DNSInfoHC
    Mock Get-PrintConfiguration
    Mock Get-Printer
    Mock Get-PrinterDriver
    Mock Get-PrinterPort
    Mock Get-PrinterSNMPInfoHC
    Mock Get-ServersHC
    Mock Invoke-Command
    Mock Send-MailHC
    Mock Write-EventLog
}

Describe 'Prerequisites' {
    Context 'send an error mail to the admin when' {
        It "the OS version is not higher than Windows Server 2012" {
            Mock Get-CimInstance {
                [PSCustomObject]@{
                    Version = '6.2.1'
                }
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like '*PrintManagement*')
            }
        }

        It 'the file ComputersNotInOU is not found' {
            $testNewParams = $testParams.Clone()
            $testNewParams.ComputersNotInOU = 'NotFound.txt'
            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*file*not found*")
            }
        }
    }
}
Describe 'the following is retrieved from each computer' {
    It 'printer queues' {
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }

        .$testScript @testParams

        $GetPrinterJobResults.ComputerName | 
        Should -Be $testGetPrinter[0].ComputerName
        $GetPrinterJobResults.Data | Should -Not -BeNullOrEmpty
    }
    It 'drivers' {
        Mock Get-PrinterDriver {
            $testGetPrinterDriver
        }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }

        .$testScript @testParams

        $PrinterDrivers.ComputerName | 
        Should -Be $testGetPrinterDriver[0].ComputerName
        $PrinterDrivers.Name | Should -Be $testGetPrinterDriver[0].Name
    }
    It 'configurations' {
        Mock Get-PrintConfiguration {
            Start-Job { $Using:testGetPrintConfiguration }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }

        .$testScript @testParams

        $PrintConfigurations.ComputerName | 
        Should -Be $testGetPrintConfiguration[0].ComputerName
        $PrintConfigurations.PrinterName | 
        Should -Be $testGetPrintConfiguration[0].PrinterName
    }
    It 'ports' {
        Mock Get-PrinterPort {
            $testGetPrinterPort
        }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }

        .$testScript @testParams

        $PrinterPorts.ComputerName | 
        Should -Contain $testGetPrinterPort[0].ComputerName
        $PrinterPorts.Name | 
        Should -Contain $testGetPrinterPort[0].Name
    }
}
Describe 'when printers cannot be retrieved or something goes wrong' {
    It 'an error is stored for that computer' {
        Mock Get-Printer {
            Start-Job { throw 'Oops' }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }

        .$testScript @testParams

        $GetPrinterJobResults.ComputerName | 
        Should -Be $testGetPrinter[0].ComputerName
        $GetPrinterJobResults.Data | Should -BeNullOrEmpty
        $GetPrinterJobResults.Error | Should -EQ 'Oops'
    }
}
Describe 'add property to PrinterQueues' {
    It 'SNMP details' {
        Mock Get-PrinterPort {
            $testGetPrinterPort
        }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }
        Mock Get-PrinterSNMPInfoHC {
            $testGetPrinterSNMPInfoHC
        }

        .$testScript @testParams

        $PrinterQueues[0].SNMP_Status | Should -Be $testGetPrinterSNMPInfoHC[0].SNMP_Status
    }
    It 'DNS info for port host address' {
        Mock Get-PrinterPort {
            $testGetPrinterPort
        }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }
        Mock Get-DNSInfoHC {
            $testGetDNSInfoHC
        }

        .$testScript @testParams

        $PrinterQueues[0].DNS_PortHostAddressToName | 
        Should -Be $testGetDNSInfoHC[1].HostName
    }
    It 'DNS info for name' {
        Mock Get-PrinterPort {
            $testGetPrinterPort
        }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }
        Mock Get-DNSInfoHC {
            $testGetDNSInfoHC
        }

        .$testScript @testParams

        $PrinterQueues[0].DNS_PrinterNameToIP | 
        Should -Be $testGetDNSInfoHC[0].IP
    }
    It 'PrinterPorts' {
        Mock Get-PrinterPort {
            $testGetPrinterPort
        }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }

        .$testScript @testParams

        $PrinterQueues[0].PortDescription | 
        Should -Be $testGetPrinterPort[0].Description
    }
    It 'PrinterDrivers' {
        Mock Get-PrinterPort {
            $testGetPrinterPort
        }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }
        Mock Get-PrinterDriver {
            $testGetPrinterDriver
        }

        .$testScript @testParams

        $PrinterQueues[0].DriverManufacturer | 
        Should -Be $testGetPrinterDriver[0].Manufacturer
    }
    It 'PrintConfiguration' {
        Mock Get-PrinterPort {
            $testGetPrinterPort
        }
        Mock Get-Printer {
            Start-Job { $Using:testGetPrinter }
        } -ParameterFilter { $AsJob -eq $true }
        Mock Get-ServersHC {
            $testGetPrinter[0].ComputerName
        }
        Mock Get-PrinterDriver {
            $testGetPrinterDriver
        }
        Mock Get-PrintConfiguration {
            Start-Job { $Using:testGetPrintConfiguration }
        } -ParameterFilter { $AsJob -eq $true }

        .$testScript @testParams

        $PrinterQueues[0].PaperSize | 
        Should -Be $testGetPrintConfiguration[0].PaperSize
    }
}
    

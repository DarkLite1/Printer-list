#Requires -Version 5.1
#Requires -Modules ImportExcel

<#
    .SYNOPSIS
        Retrieve a list of installed printers on servers.

    .DESCRIPTION
        Retrieve a list of all installed printers on servers found in the 
        specified organizational units. The list will be send by e-mail with
        an Excel file in attachment containing all printers found.

    .PARAMETER OU
        Organizational units where we search for server names.

    .PARAMETER ImportFile
        This file contains the server names that are not present in the OU's. This could be useful for Citrix servers in other OU's managed by SDO.

    .NOTES
        TROUBLESHOOTING
        Get-Printer not working on specific clients
        - No permissions
            Add user to local 'Administrator' group

        - WinRM 2.0 not installed (Win Srv 2003)
            Install it and run in CMD: 'winrm quickconfig' > Confirm 'Y'

        - RegKey missing (Win Srv 2003)
            Enter-PSSession $ComputerName
            New-ItemProperty -Path 'REGISTRY::HKLM\Software\Policies\Microsoft\Windows NT\Printers' -Name RegisterSpoolerRemoteRpcEndPoint -PropertyType DWORD -Value 1 -Verbose
            Restart-Service -Name Spooler -Verbose
            Exit-PSSession
#>

Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [Parameter(Mandatory)]
    [String[]]$OU,
    [String]$ComputersNotInOU,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Printers\Printer list\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Function Add-PropertyHC {
        Param (
            [Parameter(Mandatory)]
            [PSCustomObject[]]$InputObject,
            [Parameter(Mandatory)]
            [ScriptBlock]$Filter,
            [String]$Name
        )

        Foreach ($I in ($InputObject.Where( $Filter, 'First'))) {
            $Members = [Ordered]@{ }
            foreach ($Property in ($I.PSObject.Properties.where( {
                            ($_.Name -ne 'Name') -and
                            ($_.Name -ne 'ComputerName')
                        }
                    ))) {
                $PropName = if ($Name) {
                    $Name + $Property.Name
                }
                else {
                    $Property.Name
                }

                $Members[$PropName] = $Property.Value
            }

            $P | Add-Member -NotePropertyMembers $Members -TypeName NoteProperty
        }
    }
    Function Get-JobResultHC {
        Param (
            $Job
        )

        for ( $i = 0; $i -lt $Job.Count; $i += 2 ) {

            $ReceiveParams = @{
                ErrorVariable = 'JobError'
                ErrorAction   = 'SilentlyContinue'
            }

            [PSCustomObject]@{
                ComputerName = $Job[$i]
                State        = $Job[$i + 1].State
                Data         = $Job[$i + 1] | Receive-Job @ReceiveParams
                Error        = if ($JobError) {
                    $JobError
                    $JobError.ForEach( { $Error.Remove($_) })
                }
            }
        }
    }

    Try {
        $Error.Clear()
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        Get-Job | Remove-Job -EA Ignore

        if (
            [Version](Get-CimInstance -ClassName Win32_OperatingSystem -Verbose:$false).Version -lt [Version]'6.2.9200'
        ) {
            throw "This script is intended to run on Windows Server 2012 or later, please use an OS that supports the module 'PrintManagement'"
        }

        if (
            ($ComputersNotInOU) -and 
            (-not (Test-Path -LiteralPath $ComputersNotInOU -PathType Leaf))
        ) {
            throw "File '$ComputersNotInOU' not found that contains the computer names that are not available in the OU"
        }

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
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Get computer names
        $ServerParams = @{
            OU = $OU
        }
        if ($ComputersNotInOU) {
            $ServerParams.Path = $ComputersNotInOU
        }
        $ComputerNames = @(Get-ServersHC @ServerParams | Sort-Object -Unique)

        Write-Verbose "Retrieved $($ComputerNames.Count) computers from AD and/or the import file"

        for ($i = 0; $i -lt $ComputerNames.Count; $i++) {
            $ComputerNames[$i] = $ComputerNames[$i].toUpper()
        }
        #endregion

        #region Get printer queues
        Write-EventLog @EventVerboseParams -Message 'Get printer queues'

        $GetPrinterJobs = Foreach ($C in $ComputerNames) {
            $C
            Get-Printer -ComputerName $C -AsJob
        }

        $null = Get-Job | Wait-Job -Timeout 900 -EA Ignore
        $GetPrinterJobResults = @(Get-JobResultHC -Job $GetPrinterJobs)
        Get-Job | Remove-Job -Force

        $ComputersWithPrinters, $ComputersWithoutPrinters = $GetPrinterJobResults.where( { $_.Data }, 'Split')
        $PrinterQueues = @($ComputersWithPrinters.Data)

        Write-EventLog @EventVerboseParams -Message "A total of $($PrinterQueues.count) printer queues where found on $($ComputersWithPrinters.Count) computers:`n`n$($ComputersWithPrinters.ComputerName -join "`n")"

        Write-EventLog @EventVerboseParams -Message "$($ComputersWithoutPrinters.Count) computers have no printer queues installed or have connection issues:`n`n$($ComputersWithoutPrinters.ComputerName -join "`n")"
        #endregion

        #region Get printer drivers and ports
        <#
            Retrieving printer drivers and ports is not done in jobs. When jobs 
            are used it can happen that not all ports or drivers are retrieved.
        #>
        $PrinterDrivers = @()
        $PrinterPorts = @()

        Foreach ($Computer in $ComputersWithPrinters) {
            Write-Verbose "Computer '$($Computer.ComputerName)' get drivers"
            $PrinterDrivers += Get-PrinterDriver -ComputerName $Computer.ComputerName -EA Ignore

            Write-Verbose "Computer '$($Computer.ComputerName)' get ports"
            $PrinterPorts += Get-PrinterPort -ComputerName $Computer.ComputerName -EA Ignore
        }

        Write-EventLog @EventVerboseParams -Message "Retrieved $($PrinterDrivers.Count) printer drivers from $(@($PrinterDrivers.ComputerName | Select-Object -Unique).count) computers on a total of $($ComputersWithPrinters.Count) computers that have printers"

        Write-EventLog @EventVerboseParams -Message "Retrieved $($PrinterPorts.Count) printer ports from $(@($PrinterPorts.ComputerName | Select-Object -Unique).count) computers on a total of $($ComputersWithPrinters.Count) computers that have printers"
        #endregion

        #region Get printer configurations
        <#
            Retrieving the printer configs is done in jobs because for some 
            printers the call freezes the script. To avoid this the job runtime 
            is limited to a specific time span after which the job will be 
            stopped.

            Running multiple jobs at once failed to retrieve all details, hence 
            the use of 1 job at a time.
        #>
        $jobTimer = @{ }
        $GetPrintConfigJobs = @()

        $maxConcurrentJobs = 1
        $maxSecondsPerJob = 60

        $WaitToLaunchJob = {
            do {
                $EnumeratedJobs = @($jobTimer.GetEnumerator())
                $EnumeratedJobs.Where( {
                        ($_.Value.IsRunning) -and
                        ($_.Name.State -ne 'Running')
                    }).ForEach( {
                        $_.Value.Stop()
                    })
                $EnumeratedJobs.Where( {
                        ($_.Value.IsRunning) -and
                        ($_.Value.Elapsed.TotalSeconds -ge $maxSecondsPerJob)
                    }).Foreach( {
                        $_.Value.Stop()
                        Write-Warning "Stop job '$($_.Name.Name)' that ran for '$($_.Value.Elapsed.TotalSeconds)' seconds"
                        Stop-Job $_.Name
                    })

                $running = @(Get-Job -State Running)
                $Wait = $running.Count -ge $maxConcurrentJobs

                if ($Wait) {
                    Write-Verbose 'Waiting for jobs to finish'
                    $null = $running | Wait-Job -Any -Timeout 1
                }
            } while ($Wait)
        }

        foreach ($Printer in ($ComputersWithPrinters.Data | Sort-Object ComputerName, Name)) {
            & $WaitToLaunchJob

            Write-Verbose "Computer '$($Printer.ComputerName)' get configuration of printer '$($Printer.Name)'"
            if ($Job = $Printer | Get-PrintConfiguration -AsJob -EA Ignore) {
                $GetPrintConfigJobs += $Job
                $JobTimer[$Job] = [System.Diagnostics.Stopwatch]::StartNew()
            }
        }

        Write-Verbose "Wait for jobs 'GetPrinterConfig' to finish"
        $null = Get-Job | Wait-Job -Timeout $maxSecondsPerJob -EA Ignore

        $SelectParams = @{
            Property        = '*'
            ExcludeProperty = @('RunspaceId', 'PSShowComputerName', 'PSComputerName',
                'CimClass', 'CimInstanceProperties', 'CimSystemProperties')
        }

        $PrintConfigurations = @($GetPrintConfigJobs |
            Receive-Job -EA Ignore | Select-Object @SelectParams)
        Write-EventLog @EventVerboseParams -Message "Retrieved $($PrintConfigurations.Count) printer configurations for $(@($ComputersWithPrinters.Data).Count) printers"

        Get-Job | Remove-Job -Force
        #endregion

        #region Get SNMP details
        if ($UniquePrinterPorts = $PrinterPorts.PrinterHostAddress | Sort-Object -Unique) {
            Write-EventLog @EventVerboseParams -Message "Get printer SNMP details for $($UniquePrinterPorts.count) unique printer port host addresses"

            $SNMP = Get-PrinterSNMPInfoHC -ComputerName $UniquePrinterPorts |
            Select-Object -Property * -ExcludeProperty PSComputerName, PSSourceJobInstanceId

            Foreach ($S in $SNMP) {
                $Members = @{ }
                foreach ($Property in $S.PSObject.Properties) {
                    $Members[$Property.Name] = $Property.Value
                }

                foreach ($Port in ($PrinterPorts.Where( {
                                $_.PrinterHostAddress -eq $S.SNMP_ComputerName }))
                ) {
                    foreach ($Printer in ($PrinterQueues.Where( {
                                    ($_.ComputerName -eq $Port.ComputerName) -and
                                    ($_.PortName -eq $Port.Name)
                                }))
                    ) {
                        $Printer | Add-Member -NotePropertyMembers $Members -TypeName NoteProperty
                    }
                }
            }
        }
        #endregion

        #region Get DNS details for PrinterName and PortHostAddress
        if ($uniqueDnsNames = @($PrinterQueues.Name + $UniquePrinterPorts) | Sort-Object -Unique) {
            Write-EventLog @EventVerboseParams -Message "Get DNS details for PrinterName and PortHostAddress for $($uniqueDnsNames.count) unique names"

            foreach ($P in $PrinterQueues) {
                $Members = @{
                    DNS_PortHostAddressToName = $null
                    DNS_PrinterNameToIP       = $null
                    DNS_PrinterNameToHostname = $null
                }
                $P | Add-Member -NotePropertyMembers $Members -TypeName NoteProperty
            }

            $DNS = Get-DNSInfoHC $uniqueDnsNames |
            Select-Object -Property * -ExcludeProperty PSComputerName, PSSourceJobInstanceId

            Foreach ($D in $DNS) {
                foreach ($P in @($PrinterQueues.where( { $_.Name -eq $D.ComputerName }))) {
                    $P.DNS_PrinterNameToIP = $D.IP
                    $P.DNS_PrinterNameToHostname = $D.Hostname
                }

                foreach ($Port in ($PrinterPorts.Where( {
                                $_.PrinterHostAddress -eq $D.ComputerName }))
                ) {
                    foreach ($P in ($PrinterQueues.Where( {
                                    ($_.ComputerName -eq $Port.ComputerName) -and
                                    ($_.PortName -eq $Port.Name)
                                }))
                    ) {
                        $P.DNS_PortHostAddressToName = $D.Hostname
                    }
                }
            }
        }
        #endregion

        #region Add all properties to PrinterQueues
        Write-EventLog @EventVerboseParams -Message "Add properties Port, Driver and Config to printer queues"

        foreach ($P in $PrinterQueues) {
            if ($PrinterPorts) {
                Add-PropertyHC -InputObject $PrinterPorts -Name 'Port' -Filter {
                    ($_.Name -eq $P.PortName) -and
                    ($_.ComputerName -eq $P.ComputerName) }
            }
            if ($PrinterDrivers) {
                Add-PropertyHC -InputObject $PrinterDrivers -Name 'Driver' -Filter {
                    ($_.Name -eq $P.DriverName) -and
                    ($_.ComputerName -eq $P.ComputerName)
                }
            }
            if ($PrintConfigurations) {
                Add-PropertyHC -InputObject $PrintConfigurations -Filter {
                    ($_.PrinterName -eq $P.Name) -and
                    ($_.ComputerName -eq $P.ComputerName)
                }
            }
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message  "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {
        Write-EventLog @EventVerboseParams -Message 'Create objects containing all details'

        $Printers = $PrinterQueues | Sort-Object Name, ComputerName |
        Select-Object -Property ComputerName,
        @{Name = 'PrinterName'; Expression = { $_.Name } },
        DeviceType,
        PortName,
        PortPrinterHostAddress,
        PortDescription,
        DNS_PrinterNameToIP,
        DNS_PrinterNameToHostname,
        DNS_PortHostAddressToName,
        DriverName,
        DriverProvider,
        DriverManufacturer,
        DriverMajorVersion,
        @{Name = 'DriverVersion'; Expression = { $_.DriverDriverVersion } },
        Color,
        PaperSize,
        DuplexingMode,
        Collate,
        PrinterStatus,
        Location,
        Comment,
        Published,
        Shared,
        ShareName,
        PortPortMonitor,
        PortProtocol,
        PortPortNumber,
        PortPrinterHostIP,
        DriverConfigFile,
        DriverDataFile,
        @{Name = 'DriverHardwareID'; Expression = { @($_.DriverHardwareID)[0] } },
        DriverHelpFile,
        DriverInfPath,
        DriverOEMUrl,
        DriverPath,
        DriverPrinterEnvironment,
        DriverPrintProcessor,
        DriverType,
        SNMPEnabled,
        SNMPCommunity,
        SNMP_CommunityName,
        SNMP_Status,
        SNMP_Description,
        SNMP_CountPowerOn,
        SNMP_TonerColors,
        SNMP_MaxSpeedUnit,
        SNMP_CountUnit,
        SNMP_NICSpeedMbps,
        SNMP_Name,
        SNMP_CountTotal,
        @{Name = 'SNMP_UpTimeNIC'; Expression = {
                if ($_.SNMP_UpTimeNIC) {
                    "{0:00} Days {1:00} Hours {2:00} Mins" -f
                    $_.SNMP_UpTimeNIC.Days, $_.SNMP_UpTimeNIC.Hours, $_.SNMP_UpTimeNIC.Minutes
                }
            }
        },
        SNMP_MaxSpeed,
        SNMP_TonerNames,
        SNMP_Toners,
        SNMP_Location,
        SNMP_Model,
        SNMP_SMTP,
        SNMP_SN,
        @{Name = 'SNMP_Alert'; Expression = { $_.SNMP_Alert -join ', ' } },
        SNMP_Contact,
        Type, Datatype, DefaultJobPriority, JobCount,
        KeepPrintedJobs, PrintProcessor, Priority, StartTime, UntilTime

        $MailParams = @{ }

        $ExcelParams = @{
            Path         = $LogFile + '.xlsx'
            AutoSize     = $true
            FreezeTopRow = $true
        }
        if ($Printers) {
            Write-EventLog @EventOutParams -Message "Export to worksheet 'Printers'"

            $Params = @{
                WorkSheetName      = 'Printers'
                TableName          = 'Printers'
                NoNumberConversion = 'PortPrinterHostAddress', 'PortName', 'DriverVersion', 'DNS_PrinterNameToIP'
            }
            $Printers | Export-Excel @ExcelParams @Params

            $MailParams.Attachments = $ExcelParams.Path
        }

        $ErrorCollection = @()

        $ErrorCollection += $GetPrinterJobResults.Where( { $_.Error }) |
        Select-Object ComputerName, @{Name = 'Error'; Expression = { $_.Error -join ', ' } }

        $ErrorCollection += $Error.Exception.Message |
        Select-Object @{Name = 'ComputerName'; Expression = { $env:COMPUTERNAME } }, @{N = 'Error'; E = { $_ } }

        if ($ErrorCollection) {
            Write-EventLog @EventErrorParams -Message "$($ErrorCollection.count) errors:`n`n$($ErrorCollection[0..7].Error -join "`n")`n ..."

            $ErrorCollection | Sort-Object Error, ComputerName |
            Export-Excel @ExcelParams -WorksheetName 'Errors' -TableName 'Errors'

            $MailParams.Attachments = $ExcelParams.Path
        }

        $HtmlOu = $OU | ConvertTo-OuNameHC -OU | Sort-Object | ConvertTo-HtmlListHC -Header 'Organizational units:'

        $ComputerWithPrinterCount = ($Printers.ComputerName | Group-Object).Count
        $ComputerCount = $ComputerNames.Count
        $PrinterCount = $Printers.Count
        $ErrorCount = $ErrorCollection.Count

        $MailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = "$PrinterCount printers"
            Message   = "<p>Scanning <b>$ComputerCount computers</b> within the organizational units of active directory and in the manual import file, revealed a total of <b>$PrinterCount installed printer queues</b> on <b>$ComputerWithPrinterCount computers</b>.</p>
            $(if ($ErrorCount) {"<p>Detected <b>$ErrorCount errors</b> during execution.</p>"})
            <p><i>* Check the attachment for details</i></p>", $HtmlOu
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}
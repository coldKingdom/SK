<#
.SYNOPSIS
  Install updates for Windows 10

.DESCRIPTION
  The purpose is to install updates for Windows 10

.INPUTS
  None

.OUTPUTS
  Log file stored in $env:TEMP\UpdateLog_2019-05-16.log

.NOTES
  Version:        1.0
  Author:         Andreas Björklund
  Creation Date:  2019-05-16
  Purpose/Change: Initial script development

  Version:        1.1
  Author:         Andreas Björklund
  Creation Date:  2019-05-17
  Purpose/Change: Added parameters and some cleanup
  
.EXAMPLE
  Install-Updates.ps1 -RebootTool \\server\shutdowntool.exe -RebootToolSwitches "/t:7200 /m:1440 /r /c" -MeasureServer "server.local" UpdateSharePath-"\\server\ClientHealth$\Updates" -TemporaryDownloadPath "$env:TEMP\DownloadUpdates"
#>

#requires -RunAsAdministrator 
#requires -Version 4

[CmdLetBinding()]
Param(
    [Parameter()]
    [string]$RebootTool = "",

    [Parameter()]
    [string]$RebootToolSwitches = "",

    [Parameter()]
    [string]$MeasureServer = "",

    [Parameter()]
    [string]$UpdateSharePath = "",

    [Parameter()]
    [string]$TemporaryDownloadPath = "",

    [Parameter()]
    [string]$LogPath = [String]::Format("{0}\UpdateLog_{1}.log", $env:TEMP, (Get-Date -Format "yyyy-MM-dd")),

    [Parameter()]
    [string]$UpdateLogPath = $env:TEMP
)

# Start the logging
Start-Transcript -Path $LogPath -ErrorAction Stop

function Test-InternetConnection {
    Write-Host "Trying to ping the server $MeasureServer.."

    $Ping = (Test-NetConnection -ComputerName $MeasureServer).PingReplyDetails

    Write-Host ([String]::Format("Ping time: {0}ms", $Ping.RoundtripTime))

    if ($null -ne $Ping) {
        if ($Ping.RoundtripTime -le 50) {
            Write-Host "The beast is alive! We can ping to $MeasureServer, hurray!"

            return $true
        }
    }    

    return $false
}

function Test-SharedFolder {
    Write-Host "Alright.. Calm down a bit! So the beast is alive. What next? Can we connect to the share? Lets find out!"

    Try {
        Test-Path -Path $UpdateSharePath -IsValid
        Write-Host "We did find out this: a connection was made. Please move along!"

        return $true
    }
    Catch {
        Write-Host "No, we couldn't. Lovely! Exiting and slowly crawling my way back to the office."
    }

    return $false
}

function New-RebootTask ($TaskName = "Start - Reboot App") {
    if (Get-ScheduledTask -TaskName "Start - Reboot App" -ErrorAction SilentlyContinue) {        
        Try {
            Write-Host "Holy macarony. The task already exists. Remove it, pronto!"
            Unregister-ScheduledTask -TaskName $TaskName -Verbose
            Write-Host "Seems like it went missing, great! Lets add it, again."
        }
        Catch {
            Write-Host "Oh no, something went wrong when trying to delete the task `"$TaskName`""
            Write-Host $_.Exception.Message
        }
    }

    $task_action = New-ScheduledTaskAction -Execute $RebootTool -Argument $RebootToolSwitches
    $task_trigger = New-ScheduledTaskTrigger -Once -At "1901-01-01 12:00:00"
    $task_settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -Compatibility At
    $task_principal = New-ScheduledTaskPrincipal -GroupId "BUILTIN\Användare"

    Try {
        Write-Host "Trying to add the schedule task `"$TaskName`""
        Register-ScheduledTask -TaskName $TaskName -Action $task_action -Trigger $task_trigger -Settings $task_settings -Principal $task_principal -ErrorAction Stop
        Write-Host "It worked out pretty good."
    }
    Catch {
        Write-Host "Oh no, something went wrong when trying to register the task `"$TaskName`""
        Write-Host $_.Exception.Message
    }
}

function Start-RebootTask {
    If (Get-ScheduledTask -TaskName "ConfigMgr Client Health - Reboot on demand") {       
		Try {
			Write-Host "We already have a reboot app since like forever.. I suggest we use it."
			Start-ScheduledTask -TaskName "ConfigMgr Client Health - Reboot on demand"
			Write-Host "The task is now running!"
		}
		Catch {
			Write-Host "Something doesn't add up with the reboot task starting execution thingy.."
			Write-Host $_.Exception.Message
		}
		
    }
    Else {
        Try {
            New-RebootTask
            Start-ScheduledTask -TaskName "Start - Reboot App"
        }
        Catch {
            Write-Host "Something went terribly wrong the reboot."
            Write-Host $_.Exception.Message
        }
    }
}

#Based on <http://gallery.technet.microsoft.com/scriptcenter/Get-PendingReboot-Query-bdb79542>
function Test-RebootRequired {
    $result = @{
        CBSRebootPending            = $false
        WindowsUpdateRebootRequired = $false
        FileRenamePending           = $false
        SCCMRebootPending           = $false
    }

    #Check CBS Registry
    $key = Get-ChildItem "HKLM:Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction Ignore
    if ($null -ne $key) {
        $result.CBSRebootPending = $true
    }

    #Check Windows Update
    $key = Get-Item "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction Ignore
    if ($null -ne $key) {
        $result.WindowsUpdateRebootRequired = $true
    }

    #Check PendingFileRenameOperations
    $prop = Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction Ignore
    if ($null -ne $prop) {
        #PendingFileRenameOperations is not *must* to reboot?
        #$result.FileRenamePending = $true
    }

    #Check SCCM Client <http://gallery.technet.microsoft.com/scriptcenter/Get-PendingReboot-Query-bdb79542/view/Discussions#content>
    try { 
        $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
        $status = $util.DetermineIfRebootPending()
        if (($null -ne $status) -and $status.RebootPending) {
            $result.SCCMRebootPending = $true
        }
    }
    catch { }

    #Return Reboot required

    if ($result.ContainsValue($true)) {
        Write-Host "A reboot is needed before moving on."
    }

    return $result.ContainsValue($true)
}

function Get-OSBuild {
    $OSArchitecture = ((Get-CimInstance Win32_OperatingSystem).OSArchitecture -replace ('([^0-9])(\.*)', '')) + '-Bit'
    $OSName = ([Regex]::Match((Get-CimInstance Win32_OperatingSystem).Caption, '(W.*\s\d+)')).Value + ' ' + $OSArchitecture
    $OSBuild = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber

    switch ($OSBuild) {
        10240 { $OSName = $OSName + " 1507" }
        10586 { $OSName = $OSName + " 1511" }
        14393 { $OSName = $OSName + " 1607" }
        15063 { $OSName = $OSName + " 1703" }
        16299 { $OSName = $OSName + " 1709" }
        17134 { $OSName = $OSName + " 1803" }
        17763 { $OSName = $OSName + " 1809" }
        18362 { $OSName = $OSName + " 1903" }
        default { $OSName = $OSName + " Insider Preview" }
    }

    return $OSName
}

function Get-MissingUpdates {

    If ($UpdateSharePath.EndsWith('\') -eq $true) {
        $UpdatePath = "$UpdateSharePath{0}" -f (Get-OSBuild)
    }
    Else 
    {
        $UpdatePath = "$UpdateSharePath\{0}" -f (Get-OSBuild)
    }

    $InstalledUpdates = (Get-CimInstance -ClassName Win32_QuickFixEngineering).HotFixID

    $HotfixesNeeded = @()

    ###
    # Looking for servicing stack
    ###
    Get-ChildItem "$UpdatePath\Servicing Stack" -Filter "*.msu" | ForEach-Object {
        $KBName = ($_.Name -replace ('\b(?!(KB)+(\d+)\b)\w+') -replace "\." -replace "-").ToUpper()

        If ($InstalledUpdates -notcontains $KBName) {
            $HotfixesNeeded += @{ $KBName.ToUpper() = 
                @{
                    Name = $_.Name;
                    KB   = $KBName.ToUpper();
                    Path = $_.FullName
                }
            }

            Write-Host "Yippee ki-yay! Found an servicing stack-update with the name of $KBName. Adding it priority list!"
        }
        Else {
            Write-Host "Update $KBName is not needed"
        }
    }

    ###
    # Looking for regular CU's
    ###
    Get-ChildItem $UpdatePath -Filter "*.msu" | ForEach-Object {
        $KBName = ($_.Name -replace ('\b(?!(KB)+(\d+)\b)\w+') -replace "\." -replace "-").ToUpper()

        If ($InstalledUpdates -notcontains $KBName) {
            $HotfixesNeeded += @{ $KBName.ToUpper() = 
                @{
                    Name = $_.Name;
                    KB   = $KBName.ToUpper();
                    Path = $_.FullName
                }
            }

            Write-Host "I don't know if it's good or not. But I found an cumulative update with the name of $KBName. Putting it in the list!"
        }
        Else {
            Write-Host "Update $KBName is not needed"
        }
    }

    If ($HotfixesNeeded.Count -gt 0) {

        If (!(Test-Path $TemporaryDownloadPath)) {
            New-Item $TemporaryDownloadPath -ItemType Directory -Force
        }

        foreach ($Hotfix in $HotfixesNeeded) {
            If (!(Test-Path $UpdateLogPath)) {
                New-Item $UpdateLogPath -ItemType Directory -Force
            }

            #$HotfixName = ($Hotfix.GetEnumerator() | Select-Object Key -ExpandProperty Value).Name
            $HotfixPath = ($Hotfix.GetEnumerator() | Select-Object Key -ExpandProperty Value).Path
            $HotfixKB = ($Hotfix.GetEnumerator() | Select-Object Key -ExpandProperty Value).KB

            Try {
                Write-Host "Downloading update $HotfixKB..."
                $CopyKB = Copy-Item -Path $HotfixPath -Destination "$TemporaryDownloadPath" -Force -ErrorAction Stop -PassThru
                $HotfixLocalPath = $CopyKB.FullName
                Write-Host ([String]::Format("Update {0} downloaded to {2}. Size is {1}MB.", $HotfixKB, (($CopyKB).Length / 1MB | ForEach-Object { '{0:0.##}' -f $_ }), $CopyKB.FullName))
            }
            Catch {
                Write-Host ([String]::Format("Something went wrong during copy of {0}: {1}", $HotfixKB, $_.Exception.Message))
            }

            Write-Host ([String]::Format("Trying to install {0}.", $HotfixLocalPath))
            Write-Host ([String]::Format("Logging to {0}", "$UpdateLogPath\$($HotfixKB).log"))
            $process = Start-Process -FilePath "wusa" -ArgumentList "`"$HotfixLocalPath`" /quiet /norestart /log:`"$UpdateLogPath\$($HotfixKB).log`"" -Wait -ErrorAction SilentlyContinue -PassThru

            If ($process.ExitCode -eq 0) {
                Write-Host "Patch $HotfixKB were installed."
            }
            ElseIf ($process.ExitCode -eq 2) {
                Write-Host "Patch $HotfixKB were installed and need of a reboot."
            }
            ElseIf ($process.ExitCode -eq 2359302) {
                Write-Host "Patch $HotfixKB are already installed on the machine."
            }
        }

        If ($HotfixesNeeded.Count -gt 0) {
            If (Test-RebootRequired) {
                Write-Host "We need a reboot after installing the updates. Starting the ShutDown-tool..."
                Start-RebootTask
            }
        }
        ElseIf (Test-RebootRequired) {
            Write-Host "We need a reboot first. Starting the ShutDown-tool..."
            Start-RebootTask
        }
    }
    Else {
        Write-Host "No new updates are found. Exiting and continuing on with life. Yay!"
    }
}

function Start-CleanFolders {
    Write-Host "Do we need a cleanup? Lets take a look!"

    If (Test-Path $TemporaryDownloadPath) {
        Write-Host "Vaccuming files in $TemporaryDownloadPath"
        Remove-Item -Path "$TemporaryDownloadPath\*.msu" -Force -Verbose -ErrorAction SilentlyContinue
	
	if ((Get-ChildItem $TemporaryDownloadPath | Measure-Object).Count -eq 0) {
        	Write-Host "Removing folder $TemporaryDownloadPath since it's empty"
        	Remove-Item -Path $TemporaryDownloadPath -Force -Verbose -ErrorAction SilentlyContinue
	}
    }
    Else {
        Write-Host "The folder $TemporaryDownloadPath is not existing. Nothing to cleanup."
    }

    If (Get-ScheduledTask -TaskName "Start - Reboot App" -ErrorAction SilentlyContinue) {
        Try {
            Write-Host "Trying to remove the schedule task `"Start - Reboot App`""
            Unregister-ScheduledTask -TaskName "Start - Reboot App" -ErrorAction SilentlyContinue
            Write-Host "Jolly good, the task were removed"
        }
        Catch {
            Write-Host "Holy smokes, it didn't go as planned"
            Write-Host $_.Exception.Message
        }
    }
}

If (Test-InternetConnection) {  
    If (Test-SharedFolder) {
        If (Test-RebootRequired) {
            Start-RebootTask
        }
        Else {
            Get-MissingUpdates
        }
    }

    Start-CleanFolders
}
Else {
    Write-Host "The connection is not good enough for downloading updates. Trying later..."
}

Stop-Transcript

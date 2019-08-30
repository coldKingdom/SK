[CmdletBinding(SupportsShouldProcess=$true)]
param(
	[parameter(Mandatory=$false)] 
	[string] $Remove=$false,
    [parameter(Mandatory=$false)] 
	[string] $Reset=$false
)

Begin {
    function Get-ObjectMembers {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory=$True, ValueFromPipeline=$True)]
            [PSCustomObject]$obj
        )
        $obj | Get-Member -MemberType NoteProperty | ForEach-Object {
            $key = $_.Name
            [PSCustomObject]@{Key = $key; Value = $obj."$key"}
        }
    }

    function Download-Printers {
        $jsonUrl = 'http://edufaq.strangnas.se/wp-json/data/v1/printers'
        $jsonTemp = "$env:TEMP\printers.json"

        # Get printer from JSON-file
        Start-BitsTransfer -Source $jsonUrl -TransferType Download -Destination $jsonTemp -RetryInterval 60

        # Test if file is downloaded, otherwise exit
        if (Test-Path $jsonTemp) {   
            $content = Get-Content $jsonTemp | ConvertFrom-Json

            Remove-Item $jsonTemp -Force -ErrorAction SilentlyContinue

            return $content
        }
        else {
            exit
        }
    }

    function Remove-Printers {
        try {
            # Tar bort nätverksskrivare
            Get-Printer | Where-Object { $_.Shared -eq $true } | % {
                Write-Host "Tar bort skrivare $($_.Name)"
                Remove-Printer -Name $_.Name
            }

            # Tar bort tidigare Utskrift
            Get-Printer | Where-Object { $_.Name -eq "Utskrift" } | % {
                Write-Host "Tar bort skrivare $($_.Name)"
                Remove-Printer -Name $_.Name
                Remove-PrinterPort -Name $_.PortName
            }
        }
        catch {
            Write-Host $_.Exception.Message
        }
    }

    function Install-Printers {
        $getJsonPrinters = Download-Printers

        # Get name and explode the name into pieces
        if ($env:COMPUTERNAME -match 'EDU') {
            $computerName = [Regex]::Match($env:COMPUTERNAME, '(?=EDU)(EDU)(?:-)(\w+)')
        }
        else {
            $computerName = [Regex]::Match($env:COMPUTERNAME, '(?=SHRD|ADM)(SHRD|ADM)(?:-)([A-Za-z]{2,4})([0-9]+)')
        }

        # If the computer needs a printer that is not in the group, but explicitly named in the JSON
        if (($getJsonPrinters.computerNames | Get-ObjectMembers | Where-Object Key -Match $computerName.Groups[0] | Select-Object -First 1 | Measure-Object).Count -eq 1) {
            $printers = $getJsonPrinters.computerNames | Get-ObjectMembers | Where-Object Key -Match $computerName.Groups[0] | Select-Object -First 1
        }
        else { #otherwise take the printer from group
        
            #Match on group group level, like SHRD-AKEA*
            if (($getJsonPrinters.computerGroups | Get-ObjectMembers | Where-Object Key -Match ('{0}-{1}' -f $computerName.Groups[1], $computerName.Groups[2]) | Select-Object -First 1 | Measure-Object).Count -eq 1) {
                $printers = $getJsonPrinters.computerGroups | Get-ObjectMembers | Where-Object Key -Match ('{0}-{1}' -f $computerName.Groups[1], $computerName.Groups[2]) | Select-Object -First 1
            } #Match on top group level, like SHRD-*
            elseif (($getJsonPrinters.computerGroups | Get-ObjectMembers | Where-Object Key -Match ('{0}' -f $computerName.Groups[1]) | Select-Object -First 1 | Measure-Object).Count -eq 1) {
                $printers = $getJsonPrinters.computerGroups | Get-ObjectMembers | Where-Object Key -Match ('{0}' -f $computerName.Groups[1]) | Select-Object -First 1
            }
        }

        if ($printers -ne $null) {
            $printerServerName = $getJsonPrinters.defaultServerSettings.serverName
            $printerName = $getJsonPrinters.defaultServerSettings.printerName

            if (![String]::IsNullOrEmpty($printers.Value.serverName)) {
                $printerServerName = $printers.Value.serverName
            }

            if (![String]::IsNullOrEmpty($printers.Value.printerName)) {
                $printerName = $printers.Value.printerName
            }

            $connectionName = ('\\{0}\{1}' -f $printerServerName, $printerName)

            If (Get-Printer -Name $connectionName -ErrorAction SilentlyContinue) {
                Exit
            }
            else {
                Remove-Printers
            }
        }
        else {
            Exit
        }

        try {   
            Write-Host 'Lägger till skrivaren...'

            If (!(Get-Printer -Name $connectionName -ErrorAction SilentlyContinue)) {
                Add-Printer -ConnectionName $connectionName -ErrorAction SilentlyContinue
            }
            Else {
                Get-Printer | Where-Object { $_.ComputerName -eq $printerServerName -and $_.Comment -eq 'EDU-skrivare' } | Remove-Printer -ErrorAction SilentlyContinue

                Add-Printer -ConnectionName $connectionName -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
}
Process {
    If ($Reset -eq $true) {
        Remove-Printers
        Install-Printers
    }
    ElseIf ($Remove -eq $true) {
        Remove-Printers
    }
    Else {
        Install-Printers
    }
}
End { }

function Copy-FirewallLog {
    <#
    .SYNOPSIS
    Download Windows firewall log files from hosts remotely.
    .DESCRIPTION
    This script downloads the Windows firewall logs from remote hosts.

    In order to make it work, the remote host must enable Powershell remote management
    by using Enable-PSRemoting cmdlet.

    .PARAMETER ComputerName
    Names of hosts whose firewall logs to be downloaded. Default to be the current computer.
    .PARAMETER LogPath
    Path to Windows firewall logs.
    .PARAMETER OutputPath
    Where the firewall logs will be downloaded. Set as the current folder by default.
    .PARAMETER Credential
    Credential used to access hosts. If not assigned, the current user will be used.
    Note if UseDifferentCredentials switch is turned on, the Credential entered here
    will be ignored.
    .PARAMETER UseDifferentCredentials
    By default the cmdlet asks for credential just once. This credential will be used
    to access all hosts in the list. If you want to use different credentials for different
    hosts, turn on this switch.
    #>

    param (
        [string[]] $ComputerName = $env:COMPUTERNAME,
        [string] $LogPath = 'C:\Windows\System32\LogFiles\Firewall\pfirewall.log.old',
        [string] $OutputPath = '.\',
        [PSCredential] $Credential,
        [switch] $UseDifferentCredentials
    )

    #Confirm the output folder exists.
    if( -not (Test-Path $outputPath -PathType Container) ) {
        Write-Host 'Export folder not found.'
        return
    }
    $outputFullPath = Get-Item $outputPath | Select-Object -ExpandProperty FullName

    foreach ($Computer in $ComputerName) {
        if($useDifferentCredentials) {
            $Credential = Get-Credential
        }
        if($Credential) { #Use entered credential.
            $session = New-PSSession -ComputerName $Computer -Credential $Credential
        }
        else { #Use current user credential.
            $session = New-PSSession -ComputerName $Computer
        }

        $targetHash = (Invoke-Command -Session $session `
            -ScriptBlock { (Get-FileHash $args[0]).Hash } -ArgumentList $logpath)

        $downloadedFilePath = Join-Path -Path $outputFullPath -ChildPath "$Computer`_*.log"
        #If there's no downloaded logs, then downloads directly without comparing.
        if( Test-Path $downloadedFilePath) {
            $dowanloedLatestPath = Get-ChildItem $downloadedFilePath |
                Sort-Object LastWriteTime | 
                Select-Object -ExpandProperty FullName -Last 1
            $downloadedLatestHash = (Get-FileHash -Path $dowanloedLatestPath).Hash

            #Compare hashes of latest files. No need to download new one if they are the same.
            if($targetHash -eq $downloadedLatestHash) {
                Write-Output 'Firewall log not updated. No file is downloaded.'
                continue
            }
        }

        $timeStr = Get-Date -Format yyyyMMdd_HHmmss
        $newFilePath = Join-Path -Path $outputFullPath -ChildPath "$Computer`_$timeStr.log"
        Copy-Item -Path $logpath -Destination $newFilePath -FromSession $session
        Write-Output 'Firewall log updated. New log is downloaded.'

        Remove-PSSession -Session $session
    }
}
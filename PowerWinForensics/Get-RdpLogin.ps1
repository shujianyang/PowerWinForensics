function Get-RdpLogin {
    <#
    .SYNOPSIS
    Read Windows remote desktop log and list connections.
    .DESCRIPTION
    This function reads Windows Remote Desktop log and list all succuessful connections and
    related information, including usernames and IP addresses.
    .PARAMETER ComputerName
    Names of hosts on which remote desktop connection is allowed.
    .PARAMETER Credential
    Credential used to access hosts. If not assigned, the current user will be used.
    .PARAMETER MaxEvents
    Maximun records in the event log to read.
    #>

    param (
        [string]$ComputerName = $env:COMPUTERNAME,
        [PSCredential] $Credential,
        [int]$MaxEvents = 0
    )

    if ($Credential) {
        $session = New-PSSession -ComputerName $ComputerName -Credential $Credential
    }
    else {
        $session = New-PSSession -ComputerName $ComputerName
    }

    try {
        $ScrBlock = {
            param ([int]$eventNum = 0)
            $logName = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'
            $filterHash = @{
                LogName = $logName;
                ID = 1149
            }
            if($eventNum) {
                Get-WinEvent -FilterHashtable $filterHash -MaxEvents $eventNum
            }
            else {
                Get-WinEvent -FilterHashtable $filterHash
            }
        }

        $event = Invoke-Command -Session $session `
                    -ScriptBlock $ScrBlock `
                    -ArgumentList $MaxEvents
    }
    finally {
        Remove-PSSession -Session $session
    }

    $loginEvent = foreach ($login in $event) {
        $msg = $login.Message
        $domain = ($msg.split("`n")[-2]).split(' ')[-1]
        $user = ($msg.split("`n")[-3]).split(' ')[-1]
        $IP = ($msg.split("`n")[-1]).split(' ')[-1]
        
        $loginObj = New-Object PSObject -Property @{
            Time = $login.TimeCreated
            Domain = [string]$domain
            User = [string]$user
            IP = [string]$IP
        }
        $loginObj
    }

    $loginEvent 
}

param (
    [string]$ComputerName = $env:COMPUTERNAME,
    [int]$MaxEvents = 50
)

$session = New-PSSession -ComputerName $ComputerName

$ScrBlock = {
    param ([int]$eventNum = 0)
    $logName = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'
    $filterHash = @{
        LogName = $logName;
        ID = 1149
    }
    Get-WinEvent -FilterHashtable $filterHash -MaxEvents $eventNum
}

$event = Invoke-Command -Session $session `
            -ScriptBlock $ScrBlock `
            -ArgumentList $MaxEvents
Remove-PSSession -Session $session

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


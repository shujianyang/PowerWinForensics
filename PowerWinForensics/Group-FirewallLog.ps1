function Group-FirewallLog {
    <#
    .SYNOPSIS
    Read existing firewall logs and generate new ones grouped by different time periods.
    .DESCRIPTION
    This function generate firewall log entries and split them into different groups(files) 
    by timestamps.

    For example, the firewall logs can be split into several files. Each file contains
    entries of four-hour long.

    .PARAMETER LogPath
    Path to Windows firewall logs.
    .PARAMETER OutputPath
    Where the grouped logs will be generated. Set as the current folder by default.
    .PARAMETER Prefix
    Prefix to the generated logs.
    .PARAMETER BeginHour
    At which hour of a day the logs will be splitted.
    .PARAMETER Interval
    The length of time the grouped firewall logs contain in hours.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $LogPath,
        [string] $OutputPath = '.\',
        [string] $Prefix = 'FirewallLog_',
        [ValidateRange(0,23)]
        [int] $BeginHour = 7,
        [Validatescript({ 24 % $_ -eq 0})]
        [ValidateRange(0,24)]
        [int] $Interval = 12
    )

    [string[]]$fullLogPaths = Get-ChildItem $LogPath | Select-Object -ExpandProperty FullName

    $fileName
    $logTime
    $preLogEndTime = $null
    foreach ($log in $fullLogPaths) {
        $entries = Get-Content -Path $log | Where-Object {$_ -match '^\d{4}-\d{2}-\d{2} '}

        $logStartTime = [datetime]($entries[0].subString(0,19))
        $logEndTime = [datetime]($entries[$entries.Length-1].subString(0,19))

        $splitTimes = getSplitPoints $logStartTime $logEndTime $BeginHour $Interval

        $splitIndexs = (,0)
        foreach( $bp in $splitTimes ) {
            $splitIndexs += binarySearch $entries $bp
        }
        $splitIndexs += $entries.Length

        $pointer = 1
        [string[]]$output = $entries[0..($splitIndexs[$pointer]-1)]
        if($preLogEndTime) {
            $sameTimePeriod = inSameTimePeriod $preLogEndTime $logStartTime $BeginHour $Interval
        }
        if($preLogEndTime -and $sameTimePeriod) {
            $fileName = $Prefix + $logTime.ToString("yyyyMMdd_HHmm") + ".log"
            $filePath = Join-Path -Path $outputPath -ChildPath $fileName
            $output | Out-File -FilePath $filePath -Encoding 'ascii' -Append
            Write-Verbose "Log appended to $filePath"
        }
        else {
            $fileName = $Prefix + $logStartTime.ToString("yyyyMMdd_HHmm") + ".log"
            $filePath = Join-Path -Path $outputPath -ChildPath $fileName
            $output | Out-File -FilePath $filePath -Encoding 'ascii'
            Write-Verbose "Log written to $filePath"
        }

        if($pointer -gt ($splitIndexs.Length-2)) {continue}
        $pointer..($splitIndexs.Length-2) | ForEach-Object {
            $output = $entries[$splitIndexs[$_]..($splitIndexs[$_+1]-1)]
            if($splitTimes) {
                $logTime = $splitTimes[$_-1]
            }
            $fileName = $Prefix + $logTime.ToString("yyyyMMdd_HHmm") + ".log"
            $filePath = Join-Path -Path $outputPath -ChildPath $fileName
            $output | Out-File -FilePath $filePath -Encoding 'ascii'
            Write-Verbose "Log written to $filePath"
        }
        $preLogEndTime = $logEndTime
    }
}

#Split a time span into different periods.
function getSplitPoints {
    param(
        [datetime] $startTime,
        [datetime] $endTime,
        [int] $beginHour,
        [int] $interval
    )

    $splitPoint = $startTime.Date.AddHours($beginHour % $interval)
    while($splitPoint -le $startTime) {
        $splitPoint = $splitPoint.AddHours($interval)
    }
    while($splitPoint -lt $endTime) {
        $splitPoint
        $splitPoint = $splitPoint.AddHours($interval)
    }
}

#Identify whether two time points are in same period.
function inSameTimePeriod {
    param(
        [datetime] $earlyTime,
        [datetime] $lateTime,
        [int] $beginHour,
        [int] $interval
    )

    if(($lateTime - $earlyTime).Hours -gt $interval) { return $false }

    $splitPoint = $earlyTime.Date.AddHours($beginHour % $interval)
    while($splitPoint -le $earlyTime) {
        $splitPoint = $splitPoint.AddHours($interval)
    }
    return $splitPoint -gt $lateTime
}

#Binary search to find the index of the first entry with the given time in a log.
function binarySearch ([string[]] $array, [datetime] $time)
{
    $begin = 0
    $end = $array.Length-1
    while ($begin -le $end) {
        $mid = [int](($begin+$end)/2)
        $midtime = [datetime]($array[$mid].Substring(0,19))
        $midtimePre = [datetime]($array[$mid-1].Substring(0,19))
        if (($midtime -ge $time) -and ($midtimePre -lt $time)) {
            return $mid
        }
        if($midtime -lt $time) {
            $begin = $mid + 1
        }
        else {
            $end = $mid - 1
        }
    }
    return $mid
}

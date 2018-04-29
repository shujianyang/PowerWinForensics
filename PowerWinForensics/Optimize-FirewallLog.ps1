
function getSplitPoints {
    param(
        [datetime] $startTime,
        [datetime] $endTime,
        [int] $beginHour,
        [int] $interval
    )

    $intervalCount = [int](($startTime.Hour - $beginHour) / $interval)
    $splitTimes = $startTime.Date.AddHours($beginHour).AddHours(($intervalCount+1)*$interval)
    <#if($startTime.Hour -lt 7) {
        $splitTimes = $startTime.Date.AddHours(7)
    }
    elseIf($startTime.Hour -ge 19) {
        $splitTimes = $startTime.Date.AddDays(1).AddHours(7)
    }
    else {
        $splitTimes = $startTime.Date.AddHours(19)
    }#>
    while ($splitTimes -lt $endTime) {
        $splitTimes
        $splitTimes = $splitTimes.AddHours(12)
    }
}

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

function Optimize-FirewallLog {
    <#
    .SYNOPSIS
    Generate new firewall logs based by self-defined time periods from exsisting firewall logs.
    .DESCRIPTION
    This function generate firewall logs by the timestamps.

    For example, the firewall logs can be optimized into several files. Each one contains
    four-hour long logs.

    .PARAMETER LogPath
    Path to Windows firewall logs.
    .PARAMETER OutputPath
    Where the firewall logs will be downloaded. Set as the current folder by default.
    .PARAMETER Prefix
    Prefix to the generated logs.
    .PARAMETER BeginTime
    The hour of a day where the logs will be splitted.
    .PARAMETER Interval
    The length of time the firewall logs contain in hours.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $LogPath,
        [string] $OutputPath = '.\',
        [string] $Prefix = 'FirewallLog_',
        [int] $BeginTime = 7,
        [int] $Interval = 12
    )

    [string[]]$fullLogPaths = Get-ChildItem $LogPath | Select-Object -ExpandProperty FullName

    $fileName
    $logTime
    foreach ($log in $fullLogPaths) {
        $entries = Get-Content -Path $log | Where-Object {$_ -match '^\d{4}-\d{2}-\d{2} '}

        $logStartTime = [datetime]($entries[0].subString(0,19))
        $logEndTime = [datetime]($entries[$entries.Length-1].subString(0,19))
        if( -not $logTime) {
            $logTime = $logStartTime
            $fileName = $Prefix + $logTime.ToString("yyyyMMdd_HHmm") + ".log"
            $filePath = Join-Path -Path $outputPath -ChildPath $fileName
            if(Test-Path $filePath -PathType Leaf) {
                Remove-Item $filePath
            }
        }

        $splitTimes = getSplitPoints $logStartTime $logEndTime $BeginTime $Interval

        $indexs = (,0)
        foreach( $bp in $splitTimes ) {
            $indexs += binarySearch $entries $bp
        }
        $indexs += $entries.Length

        $pointer = 1
        [string[]]$output = $entries[0..($indexs[$pointer]-1)]
        $fileName = $Prefix + $logTime.ToString("yyyyMMdd_HHmm") + ".log"
        $filePath = Join-Path -Path $outputPath -ChildPath $fileName
        $output | Out-File -FilePath $filePath -Encoding 'ascii' -Append
        Write-Verbose "Log appended to $filePath"

        if($pointer -gt ($indexs.Length-2)) {continue}
        $pointer..($indexs.Length-2) | ForEach-Object {
            $output = $entries[$indexs[$_]..($indexs[$_+1]-1)]
            if($splitTimes) {
                $logTime = $splitTimes[$_-1]
            }
            $fileName = $Prefix + $logTime.ToString("yyyyMMdd_HHmm") + ".log"
            $filePath = Join-Path -Path $outputPath -ChildPath $fileName
            $output | Out-File -FilePath $filePath -Encoding 'ascii'
            Write-Verbose "Log re-written to $filePath"
        }
    }
}
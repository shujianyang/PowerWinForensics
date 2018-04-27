
function getBreakPoints ([datetime] $startTime, [datetime] $endTime)
{
    if($startTime.Hour -lt 7) {
        $breakpoint = $startTime.Date.AddHours(7)
    }
    elseIf($startTime.Hour -ge 19) {
        $breakpoint = $startTime.Date.AddDays(1).AddHours(7)
    }
    else {
        $breakpoint = $startTime.Date.AddHours(19)
    }
    while ($breakpoint -lt $endTime) {
        $breakpoint
        $breakpoint = $breakpoint.AddHours(12)
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
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $LogPath,
        [string] $outputPath = '.\',
        [string] $Prefix = 'FirewallLog_'
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

        $breakpoints = getBreakPoints $logStartTime $logEndTime

        $indexs = (,0)
        foreach( $bp in $breakpoints ) {
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
            if($breakpoints) {
                $logTime = $breakpoints[$_-1]
            }
            $fileName = $Prefix + $logTime.ToString("yyyyMMdd_HHmm") + ".log"
            $filePath = Join-Path -Path $outputPath -ChildPath $fileName
            $output | Out-File -FilePath $filePath -Encoding 'ascii'
            Write-Verbose "Log re-written to $filePath"
        }
    }
}
function Get-ServiceStats {
    <#
    .SYNOPSIS
    Read a series of Windows firewall logs and generate service (port) statistics.
    .DESCRIPTION
    Read a series of Windows firewall logs and generate service (port) statistics.

    .PARAMETER LogPath
    Path to logs. Multiple paths or wildcards are allowed.
    .PARAMETER MaxCount
    The number of ports with highest counts to be listed.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $LogPath,
        [int] $MaxCount = $null
    )

    # Test if log path and output path are valid.
    if( -not (Test-Path $LogPath -PathType Leaf) ) {
        Write-Host 'Log file not found.'
        return
    }

    [string[]]$fullLogPaths = Get-ChildItem $LogPath -File | Select-Object -ExpandProperty FullName

    $PortCount = @{}

    foreach ($log in $fullLogPaths) {
        #Filter out log headers
        $entries = Get-Content -Path $log | Where-Object {$_ -match '^\d{4}-\d{2}-\d{2} '}

        ForEach ($Line in $entries)
        {
            $Data = $line.split(' ')
            #$Date = $Data[0]
            #$Time = $Data[1]
            #$Action = $Data[2]
            $Protocol = $Data[3]
            $DstIp = $Data[5]
            #$SrcPort = $Data[6]
            $DstPort = $Data[7]
            $Path = $Data[16]
            
            if($Protocol -ne 'TCP' -and $Protocol -ne 'UDP') {
                continue
            }

            $Port = $Protocol, $DstPort -join ' '

            if($DstIp -ne '127.0.0.1' -and $Path -eq 'Receive') {
                $PortCount[$Port] += 1
            } 
        }
        Write-Verbose 'Finish reading $log'
    }

    $PortList = $PortCount.GetEnumerator() | Sort-Object -Property Value -Descending

    $PortStats = foreach ($Port in $PortList) {
        $p = -split $Port.Name
        $PortObj = New-Object PSObject
        Add-Member -InputObject $PortObj -MemberType NoteProperty -name 'Protocol' -value $p[0]
        Add-Member -InputObject $PortObj -MemberType NoteProperty -name 'Port' -value $p[1]
        Add-Member -InputObject $PortObj -MemberType NoteProperty -name 'Couont' -value $Port.Value
        $PortObj
    }
    $PortStats
}

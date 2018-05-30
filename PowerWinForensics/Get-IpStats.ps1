function Get-IpStats {
    <#
    .SYNOPSIS
    To read a series of Windows firewall logs and generate IP address statistics.
    .DESCRIPTION
    This function parses the Windows firewall logs and collect information about date, time, ports,
    IP addresses.

    .PARAMETER LogPath
    Path to logs. Multiple paths or wildcards are allowed.
    .PARAMETER Port
    Port(s) to analyze. All other ports will be ignored.
    The port will always be the destination port, regarless whether it is inbound or outbound.
    .PARAMETER Direction
    The direction of traffic to analyze. Accepatable values are Inbound, Outbound or All.
    Default value is Inbound, meaning only inbound traffic will be analyzed.
    .PARAMETER MaxCount
    The number of IPs with highest counts to be listed.
    .PARAMETER DisableIPLocation
    Disable IP location query so all external IP addresses will be diplayed as (external).
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $LogPath,
        [string[]] $Port,
        [ValidateSet('Inbound','Outbound','All')]
        [string] $Direction = 'Inbound',
        [int] $MaxCount = $null,
        [switch] $DisableIPLocation
    )

    # Test if log path and output path are valid.
    if( -not (Test-Path $LogPath -PathType Leaf) ) {
        Write-Host 'Log file not found.'
        return
    }

    [string[]]$fullLogPaths = Get-ChildItem $LogPath -File | Select-Object -ExpandProperty FullName

    $IpCount = @{}

    foreach ($log in $fullLogPaths) {
        #Filter out log headers
        $entries = Get-Content -Path $log | Where-Object {$_ -match '^\d{4}-\d{2}-\d{2} '}

        ForEach ($Line in $entries)
        {
            $Data = $line.split(' ')
            #$Date = $Data[0]
            #$Time = $Data[1]
            #$Action = $Data[2]
            #$Protocol = $Data[3]
            $SrcIp = $Data[4]
            #$DstIp = $Data[5]
            #$SrcPort = $Data[6]
            $DstPort = $Data[7]
            $Path = $Data[16]

            if( $Port -and ($DstPort -notin $Port) ) {
                continue
            }
            if($Direction -eq 'Inbound' -and $Path -ne 'Receive') {
                continue
            } 
            if($Direction -eq 'Outbound' -and $Path -ne 'Send') {
                continue
            } 

            $IpCount[$SrcIp] += 1
        }
        Write-Verbose "Finish reading $log"
    }

    Write-Verbose "Finding IP address location..."
    $IpList = $IpCount.GetEnumerator()
    if ($MaxCount) {
        $IpList = $IpList | Sort-Object -Property Value -Descending | Select-Object -First $MaxCount
    }
    else {
        $IpList = $IpList | Sort-Object -Property Value -Descending
    }
    $IpStats = ForEach ($IP in $IpList)
    {
        $ipAddress = $IP.Name
        $country = '(internal)'
        $region = ''
        $city = ''
        if( -not (isPrivateAddress $ipAddress)) {
            if($DisableIPLocation) {
                $country = '(external)'
            }
            else {
                $url = "http://freegeoip.net/json/$ipAddress"
                $location = Invoke-RestMethod -Method GET -Uri $url
                if ($location) {
                    $country = $location.country_name
                    $region = $location.region_name
                    $city = $location.city
                }
            }
        }

        $IpObj = New-Object PSObject
        Add-Member -InputObject $IpObj -MemberType NoteProperty -name 'IP' -value $ipAddress
        Add-Member -InputObject $IpObj -MemberType NoteProperty -name 'Country' -value $country
        Add-Member -InputObject $IpObj -MemberType NoteProperty -name 'Region' -value $region
        Add-Member -InputObject $IpObj -MemberType NoteProperty -name 'City' -value $city
        Add-Member -InputObject $IpObj -MemberType NoteProperty -name 'Count' -value $IP.value
        $IpObj
    }

    $IpStats
}

function isPrivateAddress([string] $ip)
{
    $octets = [int[]]$ip.Split('.')
    $first = $octets[0]
    $second = $octets[1]
    return ($first -eq 10) -or `
            (($first -eq 172) -and ($second -in 16..31)) -or `
            (($first -eq 192) -and ($second -eq 168))
}
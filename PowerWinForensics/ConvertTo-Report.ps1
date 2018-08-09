function ConvertTo-IPReport {
    <#
    .SYNOPSIS
    To read a series of Windows firewall logs and generate reports in Excel.
    .DESCRIPTION
    This script parses the Windows firewall logs and collect information about date, time, ports,
    IP addresses. The generated report contains detailed analysis for each log and an overall summary.

    It is recommended to get the logs organized before utilizing this script.

    The computer running this scipt must have Microsoft Office installed in order to generate report in Excel.
    .PARAMETER LogPath
    Path to logs. Multiple paths or wildcards are allowed.
    .PARAMETER OutputPath
    Where the report will be generated. By default it is the current folder.
    .PARAMETER OutputName
    File name of generated report, excluding file extension. Default value is "Report".
    .PARAMETER Port
    Port(s) to analyze. All other ports will be ignored.
    The port will always be the destination port, regarless whether it is inbound or outbound.
    .PARAMETER Direction
    The direction of traffic to analyze. Accepatable values are Inbound, Outbound or All.
    Default value is Inbound, meaning only inbound traffic will be analyzed.
    .PARAMETER DisableIPLocation
    Disable IP location query so all external IP addresses will be diplayed as (external).
    #>

    param (
        [Parameter(Mandatory = $true)]
        [string[]] $LogPath,
        [string] $OutputPath = '.\',
        [string] $OutputName = 'Report',
        [string[]] $Port,
        [ValidateSet('Inbound','Outbound','All')]
        [string] $Direction = 'Inbound',
        [switch] $DisableIPLocation
    )

    # Test if log path and output path are valid.
    if( -not (Test-Path $LogPath -PathType Leaf) ) {
        Write-Host 'Log file not found.'
        return
    }
    if( -not (Test-Path $outputPath -PathType Container) ) {
        Write-Host 'Export folder not found.'
        return
    }

    [string[]]$fullLogPaths = Get-ChildItem $LogPath -File | Select-Object -ExpandProperty FullName
    $outputFullPath = Get-Item $outputPath | Select-Object -ExpandProperty FullName
    $filePath = Join-Path -Path $outputFullPath -ChildPath ($outputName + '.xlsx')
    $suffix = 1
    while(Test-Path $filePath -PathType Leaf) {
        $outputName += "_$suffix"
        $filePath = Join-Path -Path $outputFullPath -ChildPath ($outputName+'.xlsx')
        $suffix++
    }

    $excel = New-Object -comobject Excel.Application
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $summarySheet = $workbook.Worksheets.Item(1)
    $summarySheet.Name = 'Summary'

    $summary = @()

    foreach ($log in $fullLogPaths) {
        #Filter out log headers
        $entries = Get-Content -Path $log | Where-Object {$_ -match '^\d{4}-\d{2}-\d{2} '}

        $IpCount = @{}

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
        Write-Host "Finish reading $log"
        $startTime = [datetime]($entries[0].subString(0,19))
        $endTime = [datetime]($entries[-1].subString(0,19))

        #Determin whether it is day or night
        if (($startTime.Hour -lt 7) -or ($startTime.Hour -ge 19)) {
            $DON = 'night'
        }
        else {
            $DON = 'day'
        }
        $startStr = $startTime.ToString('MMM dd ') + $DON

        $IpList = $IpCount.GetEnumerator()
        $IpOutput = ForEach ($IP in $IpList)
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

        $IpList = $IpOutput | Sort-Object -Property Count -Descending
        if($IpList.Count -gt 10) {
            $other = $IpList[10..($IpList.Length-1)] 
            $otherCount = ($other | Measure-Object -Property Count -Sum).Sum
        }

        $lastsheet = $workbook.Worksheets.Item($workbook.Worksheets.Count)
        $worksheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastsheet)
        $worksheet.Name = $startStr

        $worksheet.Cells.Item(1,1) = "Server Connection Counts on Port [$Port]"
        $worksheet.Range("A1:F1").Merge()
        $worksheet.Cells.Item(2,1) = "[$startTime - $endTime]"
        $worksheet.Range("A2:F2").Merge()
        $worksheet.Range("A1:F2").Style = "Title"

        $row = 4
        $worksheet.Cells.Item($row,1) = 'Source IP'
        $worksheet.Cells.Item($row,2) = 'Country'
        $worksheet.Cells.Item($row,3) = 'Region'
        $worksheet.Cells.Item($row,4) = 'City'
        $worksheet.Cells.Item($row,5) = 'Connection Counts'
        1..5 | ForEach-Object { $worksheet.Cells.Item($row,$_).Font.Bold = $True}
        $start = $row
        $row++

        $internalCount = 0
        $externalCount = 0
        foreach ($ip in $IpList)
        {
            $worksheet.Cells.Item($row,1) = $ip.IP
            $worksheet.Cells.Item($row,2) = $ip.Country
            $worksheet.Cells.Item($row,3) = $ip.Region
            $worksheet.Cells.Item($row,4) = $ip.City
            $worksheet.Cells.Item($row,5) = $ip.Count
            $row++

            if( isPrivateAddress $ip.IP ) {
                $internalCount += $ip.Count
            }
            else {
                $externalCount += $ip.Count
            }
        }

        $worksheet.Cells.Item($row + 1, 1) = 'Top 10'
        $worksheet.Cells.Item($row + 1, 5) = ($IpList | Measure-Object -Property Count -Sum).Sum - $otherCount
        $worksheet.Cells.Item($row + 2, 1) = 'other'
        $worksheet.Cells.Item($row + 2, 5) = $otherCount
        $worksheet.Cells.Item($row + 4, 1) = 'total internal'
        $worksheet.Cells.Item($row + 4, 5) = $internalCount
        $worksheet.Cells.Item($row + 5, 1) = 'total external'
        $worksheet.Cells.Item($row + 5, 5) = $externalCount


        $worksheet.Columns.Item("A:E").EntireColumn.AutoFit() | out-null
        $worksheet.Columns.item("C:C").columnwidth=15
        $worksheet.Columns.item("F:F").columnwidth=12

        $xlChartType = [Microsoft.Office.Interop.Excel.XLChartType]

        $objCharts = $worksheet.ChartObjects()
        $objChart = $objCharts.Add(500, 50, 500, 300)
        $pieStart = $row + 4
        $pieEnd = $row + 5
        $dataRange = $worksheet.range("A$pieStart`:A$pieEnd, E$pieStart`:E$pieEnd")
        $objChart.Chart.SetSourceData($dataRange, 2)
        #$objChart.Chart.ChartType = 70
        $objChart.Chart.ChartType = [Microsoft.Office.Interop.Excel.XLChartType]::xl3DPieExploded
        $objChart.Chart.ApplyDataLabels(5)
        $objChart.Chart.HasTitle = $true
        $objChart.Chart.ChartTitle.Text = 'Internal and External Connections Ratio'

        $objChart = $objCharts.Add(500, 400, 500, 300)
        $end = $start + 10
        $otherRow = $row + 2
        $dataRange = $worksheet.range("A$start`:A$end, E$start`:E$end, A$otherRow`:A$otherRow, E$otherRow`:E$otherRow")
        $objChart.Chart.SetSourceData($dataRange, 2)
        #$objChart.Chart.ChartType = 70
        $objChart.Chart.ChartType = [Microsoft.Office.Interop.Excel.XLChartType]::xl3DPieExploded
        $objChart.Chart.ApplyDataLabels(5)
        $objChart.Chart.HasTitle = $true
        $objChart.Chart.ChartTitle.Text = 'Top 10 Sources of Connections'

        $timeStr = $startTime.ToString('ddd') + ' ' + $DON
        $summary += ,($startTime.ToString('MMM dd yyyy'), 
                        $timeStr,
                        $internalCount, $externalCount)
    }

    $summarySheet.Cells.Item(1,1) = "Summary of Port [$Port]"
    $summarySheet.Range("A1:F1").Merge()
    $summarySheet.Range("A1:F1").Style = "Title"

    $row = 3
    $rangeTop = $row
    $summarysheet.Cells.Item($row, 1) = 'Date'
    $summarysheet.Cells.Item($row, 2) = 'Time'
    $summarysheet.Cells.Item($row, 3) = 'Internal'
    $summarysheet.Cells.Item($row, 4) = 'External'
    1..4 | ForEach-Object { $summarysheet.Cells.Item($row,$_).Font.Bold = $True }
    $row++

    foreach ($s in $summary) {
        $summarysheet.Cells.Item($row, 1) = $s[0]
        $summarysheet.Cells.Item($row, 2) = $s[1]
        $summarysheet.Cells.Item($row, 3) = $s[2]
        $summarysheet.Cells.Item($row, 4) = $s[3]
        $row++
    }
    $rangeBottom = $row - 1
    $summarySheet.Columns.item("A:A").columnwidth=14
    $summarySheet.Columns.item("B:B").columnwidth=11
    $summarySheet.Columns.item("C:E").columnwidth=9
    $summarySheet.Columns.item("F:F").columnwidth=20

    $objCharts = $summarysheet.ChartObjects()
    $dataRange = $summarysheet.range("B$rangeTop`:D$rangeBottom")
    $objChart = $objCharts.Add(450, 40, 500, 250)
    $objChart.Chart.SetSourceData($dataRange, 2)
    #$objChart.Chart.ChartType = 4
    $objChart.Chart.ChartType = [Microsoft.Office.Interop.Excel.XLChartType]::xlLine
    $objChart.Chart.ApplyDataLabels(5)
    $objChart.Chart.HasTitle = $true
    $objChart.Chart.ChartTitle.Text = 'Connection Counts'

    $objChart = $objCharts.Add(450, 330, 500, 250)
    $objChart.Chart.SetSourceData($dataRange, 2)
    #$objChart.Chart.ChartType = 52
    $objChart.Chart.ChartType = [Microsoft.Office.Interop.Excel.XLChartType]::xlColumnStacked
    $objChart.Chart.ApplyDataLabels(5)
    $objChart.Chart.HasTitle = $true
    $objChart.Chart.ChartTitle.Text = 'Total Connection Counts'

    $workbook.Worksheets.Item(1).Activate()
    $workbook.SaveAs($filePath)
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
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
<#
NAME: SiteScopeParser.ps1
AUTHOR: Austin Vargason
DESCRIPTION: Parses Daily Site Scope Report from XML
#>

function Get-ParsedData {
    param(
        [Parameter(Mandatory=$true)]
        [String]$FilePath
    )

    #get the xml data from the filepath provided
    $xml = [xml] (Get-Content -Path $FilePath)

    #store the report in a variable
    $report = $xml.report

    #store a variable to represent the summary data
    $summary = $report.summary

    #get the uptime sumnmary
    $summaryRow = $summary.row

    $uptimeSummary = Get-ParsedUptimeSummary($summaryRow)

    #get the Grouped data for rows
    $rowResult = Get-ParsedRowData -RowData $report.rows.row

    #get the parsed CPU data
    $availability = Get-ParsedGroupData -RowData $rowResult -LabelName "Ping Monitor on (?<ServerName>.+) Round Trip Time"

    $availabilitySummary = Get-SummarizedPingData -InputData $availability


    #get the filepath to export data
    $outputFilePath = "Output\ParsedSiteScopeReportMonthlyAvailability_" + (Get-Date -Format "MM_dd_yyyy_hh_mm_ss") + ".xlsx"

    # export the data to an excel file
    $uptimeSummary | Export-Excel -Path $outputFilePath -WorksheetName "Uptime Summary" -TableName "UptimeSummaryTable" -AutoSize
    $availabilitySummary | Export-Excel -Path $outputFilePath -WorksheetName "Availability" -TableName "AvailabilityTable" -AutoSize

}

# Function for getting the parsed uptime summary
function Get-ParsedUptimeSummary {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object[]]$SummaryRow
    )

    # array to store the result
    $result = @()

    # loop through the summary data and add into the object array
    for ($i = 0; $i -lt $SummaryRow.Count; $i++) {
        $obj = New-Object -TypeName psobject

        $rowProperties = $SummaryRow[$i] | Get-Member -MemberType Property | Select-Object -ExpandProperty Name

        foreach ($property in $rowProperties) {
            $obj | Add-Member -Name "$property" -Value ($SummaryRow[$i].$property) -MemberType NoteProperty
        }

        $result += $obj
    }

    #use custom selection to modify our property names
    $result = $result |
    Select-Object @{Label="Name";Expression={$_.name}}, @{Label="Uptime %";Expression={$_.uptime}}, @{Label="Error %";Expression={$_.error}}, @{Label="Warning %";Expression={$_.warning}}, @{Label="Last";Expression={$_.last}}

    # return the result
    return $result 
}


# function to get the group data based on the property label
function Get-ParsedGroupData {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object[]]$RowData,
        [Parameter(Mandatory=$true)]
        [String]$LabelName
    )

    $data = $RowData | Where-Object {$_.Name -match $LabelName}

    $result = @()

    foreach ($row in $data) {

        $row.Name -match $LabelName | Out-Null

        $serverName = $Matches.ServerName

        if (!$serverName.Contains(' ')) {
            $obj = New-Object -TypeName psobject

            $obj | Add-Member -Name "Server Name" -Value $serverName -MemberType NoteProperty

            foreach ($value in $row.Group) {
                $obj | Add-Member -Name $value.DateTime -Value $value.value -MemberType NoteProperty
            }

            $result += $obj
        }
    }

    $result
}

# function to parse the row data into grouped data based on the property label
function Get-ParsedRowData {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object[]]$RowData
    )

    $result = @()

    for ($i = 0; $i -lt $RowData.Count; $i++) {
        $sampleList = @()
        $dateTime = $RowData[$i].date

        $sampleData = $RowData[$i].sample

        for ($j = 0; $j -lt $sampleData.Count; $j++) {
            $sampleRowObj = New-Object -TypeName psobject

            $sampleRowMembers = $sampleData[$j] | Get-Member -MemberType Property | Select-Object -ExpandProperty Name

            foreach ($property in $sampleRowMembers) {
                $sampleRowObj | Add-Member -Name $property -Value $sampleData[$j].$property -MemberType NoteProperty
            }

            $sampleRowObj | Add-Member -Name "DateTime" -Value $dateTime -MemberType NoteProperty

            $sampleList += $sampleRowObj
        }
        
        $result += $sampleList
    }

    $result | Group-Object -Property label

}

# Function to get the summarized memory data from the data grabbed from the physical memory monitor
Function Get-SummarizedPingData {

    param (
        [Parameter(Mandatory=$true)]
        [psobject[]]$InputData
    )

    $summary = @()

    foreach ($val in $InputData) {

        $obj = New-Object -TypeName psobject

        $obj | Add-Member -Name "Server Name" -Value $val."Server Name" -MemberType NoteProperty
        $obj | Add-Member -Name "Avg" -Value 0 -MemberType NoteProperty
        $obj | Add-Member -Name "Max" -Value 0 -MemberType NoteProperty

        $sum = 0
        $max = 0
        $count = 0

        $props = $val.psobject.properties | Where-Object  {$_.Name -ne "Server Name"} | Select-Object -ExpandProperty Name


        foreach ($prop in $props) {

            $amount = $val.$prop.Replace("%", "")

            if ($amount -ne "no data" -and $amount -ne "n/a") {

                $sum += $amount

                if ($amount -gt $max) {
                    $max = $amount
                }

                $count++

            }

            $obj | Add-Member -Name $prop -Value $amount -MemberType NoteProperty
        }


        if ($count -ne 0) {
            $avg = $sum / $count

            $obj.Avg = $avg
            $obj.Max = $max
        }
        else {
            $obj.Avg = "n/a"
            $obj.Max = "n/a"
        }

        $summary += $obj
    }

    return $summary
}





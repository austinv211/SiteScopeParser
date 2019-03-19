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

    #get the measurement summary
    $measurementData = $summary.measurement

    $measurementSummary = Get-ParsedMeasurementSummary($measurementData)

    #get the error summary
    $errorData = $report.errorTimeSummary.row

    $errorTimeSummary = Get-ParsedErrorTimeSummary($errorData)

    #get the Grouped data for rows
    $rowResult = Get-ParsedRowData -RowData $report.rows.row

    #get the parsed CPU data
    $cpuResult = Get-ParsedGroupData -RowData $rowResult -LabelName "CPU Utilization Monitor on (?<ServerName>.+) Utilization"

    # # get cpu Data summary
    $cpuSummary = Get-SummarizedCPUData -CPUData $cpuResult

    #get Memory Data
    $pm_pm_used_per = Get-ParsedGroupData -RowData $rowResult -LabelName "Physical Memory Monitor on (?<ServerName>.+) Physical Memory Used %"

    # add avg to Memory Used
    $physicalMemSummary = Get-SummarizedMemoryData -MemoryData $pm_pm_used_per


    #get diskspace details
    $canGet = $true
    $i = 1
    $diskData = @()

    while ($canGet -eq $true) {
        $searchString = "Diskspace Utilization Monitor on (?<ServerName>.+) BrowsableValue" + $i

        $test_res = Get-ParsedGroupData -RowData $rowResult -LabelName $searchString
        

        if ($null -ne $test_res) {
            $test_res | Add-Member -Name DirName -Value "BrowsableValue$i" -MemberType NoteProperty
            $diskData += $test_res
        }
        else {
            $canGet = $false
        }

        $i++

    }


    # summarize the disk data
    $browsSummary = Get-SummarizedDiskData -DiskData $diskData


    #get the filepath to export data
    $outputFilePath = "ParsedSiteScopeReportMonthly_" + (Get-Date -Format "MM_dd_yyyy_hh_mm_ss") + ".xlsx"

    # export the data to an excel file
    $uptimeSummary | Export-Excel -Path $outputFilePath -WorksheetName "Uptime Summary" -TableName "UptimeSummaryTable" -AutoSize
    $measurementSummary | Export-Excel -Path $outputFilePath -WorksheetName "Measurement Summary" -TableName "MeasureSummaryTable" -AutoSize
    $errorTimeSummary | Export-Excel -Path $outputFilePath -WorksheetName "Error Time Summary" -TableName "ErrorTimeSummaryTable" -AutoSize
    $cpuSummary | Export-Excel -Path $outputFilePath -WorksheetName "CPU" -TableName "CPUTable" -AutoSize
    $physicalMemSummary | Export-Excel -Path $outputFilePath -WorksheetName "Memory" -TableName "MemoryTable" -AutoSize
    $browsSummary | Export-Excel -Path $outputFilePath -WorksheetName "File System" -TableName "DiskSpaceTable" -AutoSize

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

# function for getting the parsed measurement summary
function Get-ParsedMeasurementSummary {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object[]]$MeasurementData
    )

    $result = @()

    for ($i = 0; $i -lt $MeasurementData.Count; $i++) {
        $obj = New-Object -TypeName psobject

        $rowProperties = $MeasurementData[$i] | Get-Member -MemberType Property | Select-Object -ExpandProperty Name

        foreach ($property in $rowProperties) {
            $obj | Add-Member -Name "$property" -Value ($MeasurementData[$i].$property) -MemberType NoteProperty
        }

        $result += $obj
    }

    return $result | 
    Select-Object @{Label="Name";Expression={$_.monitor}}, @{Label="Measurement";Expression={$_.label}}, @{Label="Max";Expression={$_.max}}, @{Label="Avg";Expression={$_.ave}}, @{Label="Last";Expression={$_.last}}
}

#function for getting the parsed error time summary
function Get-ParsedErrorTimeSummary {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object[]]$ErrorData
    )

    $result = @()

    for ($i = 0; $i -lt $ErrorData.Count; $i++) {
        $obj = New-Object -TypeName psobject

        $rowProperties = $ErrorData[$i] | Get-Member -MemberType Property | Select-Object -ExpandProperty Name

        foreach ($property in $rowProperties) {
            $obj | Add-Member -Name "$property" -Value ($ErrorData[$i].$property) -MemberType NoteProperty
        }

        $result += $obj
    }

    return $result | 
    Select-Object @{Label="Name";Expression={$_.name}}, @{Label="Time in Error";Expression={$_.errorTime}}
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

# function to get the summarized data from the disk utilization monitor
Function Get-SummarizedDiskData {
    param (
        [Parameter(Mandatory=$true)]
        [psobject[][]]$DiskData
    )

    # combine browsable values
    $browsSummary = @()
    $i = 1;

    foreach ($dataList in $DiskData) {
        foreach($val in $dataList) {
            $obj = New-Object -TypeName psobject

            $count = 0
            $sum = 0
            $max = 0
            
            $props = $val.psobject.properties | Where-Object {$_.Name -ne "Server Name" -and $_.Name -ne "DirName"} | Select-Object -ExpandProperty Name
    
            $obj | Add-Member -Name "Server Name" -Value $val."Server Name" -MemberType NoteProperty
            $obj | Add-Member -Name "DirName" -Value $val.DirName -MemberType NoteProperty
            $obj | Add-Member -Name "Avg" -Value 0 -MemberType NoteProperty
            $obj | Add-Member -Name "Max" -Value 0 -MemberType NoteProperty
    
            foreach ($prop in $props) {
                $amount = $val.$prop.Replace("%", "")
    
                if ($amount -ne "no data" -and $amount -ne "n/a") {
                    $sum += $amount
    
                    if ($amount -gt $max) {
                        $max = $amount
                    }
    
                    $count++
                }
    
                $obj | Add-Member -Name $prop $val.$prop -MemberType NoteProperty
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
    
            $browsSummary += $obj
        }

        $i++
    }

    return $browsSummary
}

# Function to get the summarized memory data from the data grabbed from the physical memory monitor
Function Get-SummarizedMemoryData {

    param (
        [Parameter(Mandatory=$true)]
        [psobject[]]$MemoryData
    )

    $physicalMemSummary = @()

    foreach ($val in $MemoryData) {

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

        $physicalMemSummary += $obj
    }

    return $physicalMemSummary
}

#function to get the summarized cpu data from the data grabbed from the cpu utlization monitor
function Get-SummarizedCPUData {
    
    param (
        [Parameter(Mandatory=$true)]
        [psobject[]]$CPUData
    )

    $cpuSummary = @()

    foreach ($val in $CPUData) {
        
        $count = 0
        $sum = 0
        $max = 0

        $obj = New-Object -TypeName psobject

        $obj | Add-Member -Name "Server Name" -Value $val."Server Name" -MemberType NoteProperty
        $obj | Add-Member -Name "Avg" -Value 0 -MemberType NoteProperty
        $obj | Add-Member -Name "Max" -Value 0 -MemberType NoteProperty

        $props = $val.psobject.properties | Where-Object {$_.Name -ne "Server Name"} | Select-Object -ExpandProperty Name

        foreach ($prop in $props) {
            $amount = $val.$prop.Replace("%", "")

            if ($amount -ne "no data") {
                $sum += $amount
                $count++

                if ($amount -gt $max) {
                    $max = $amount
                }
            }

            $obj | Add-Member -Name $prop -Value $amount -MemberType NoteProperty
        }

        $avg = $sum / $count

        $obj.Avg = $avg
        $obj.Max = $max

        $cpuSummary += $obj

    }

    return $cpuSummary
    
}

Get-ParsedData .\MonthlyReport.xml


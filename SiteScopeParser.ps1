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

    #get Parsed Physical Memory Details
    $pm_virtual_mem_or_swap_space_used_per = Get-ParsedGroupData -RowData $rowResult -LabelName "Physical Memory Monitor on (?<ServerName>.+) Virtual Memory Or Swap Space Used %"

    $pm_virtual_mem_or_swap_space_mb_free = Get-ParsedGroupData -RowData $rowResult -LabelName "Physical Memory Monitor on (?<ServerName>.+) Virtual Memory Or Swap Space MB Free"

    $pm_mb_free = Get-ParsedGroupData -RowData $rowResult -LabelName "Physical Memory Monitor on (?<ServerName>.+) MB Free"

    $pm_per_used = Get-ParsedGroupData -RowData $rowResult -LabelName "Physical Memory Monitor on (?<ServerName>.+) Percent Used"

    $pm_pm_mb_free = Get-ParsedGroupData -RowData $rowResult -LabelName "Physical Memory Monitor on (?<ServerName>.+) Physical Memory MB Free"

    $pm_pm_used_per = Get-ParsedGroupData -RowData $rowResult -LabelName "Physical Memory Monitor on (?<ServerName>.+) Physical Memory Used %"

    $pm_pages_sec = Get-ParsedGroupData -RowData $rowResult -LabelName "Physical Memory Monitor on (?<ServerName>.+) Pages/sec"

    #get ping details
    $pingDetails = Get-ParsedGroupData -RowData $rowResult -LabelName "Ping Monitor on (?<ServerName>.+) Round Trip Time"

    #get virtual memory details
    $vm_virtual_mem_or_swap_space_used_per = Get-ParsedGroupData -RowData $rowResult -LabelName "Virtual Memory Utilization on (?<ServerName>.+) Virtual Memory Or Swap Space Used %"

    $vm_virtual_mem_or_swap_space_mb_free = Get-ParsedGroupData -RowData $rowResult -LabelName "Virtual Memory Utilization on (?<ServerName>.+) Virtual Memory Or Swap Space MB Free"

    $vm_mb_free = Get-ParsedGroupData -RowData $rowResult -LabelName "Virtual Memory Utilization on (?<ServerName>.+) MB Free"

    $vm_per_used = Get-ParsedGroupData -RowData $rowResult -LabelName "Virtual Memory Utilization on (?<ServerName>.+) Percent Used"

    $vm_pm_mb_free = Get-ParsedGroupData -RowData $rowResult -LabelName "Virtual Memory Utilization on (?<ServerName>.+) Physical Memory MB Free"

    $vm_pm_used_per = Get-ParsedGroupData -RowData $rowResult -LabelName "Virtual Memory Utilization on (?<ServerName>.+) Physical Memory Used %"

    $vm_pages_sec = Get-ParsedGroupData -RowData $rowResult -LabelName "Virtual Memory Utilization on (?<ServerName>.+) Pages/sec"

    #get diskspace details
    $dp_1 = Get-ParsedGroupData -RowData $rowResult -LabelName "Diskspace Utilization Monitor on (?<ServerName>.+) BrowsableValue1"

    $dp_2 = Get-ParsedGroupData -RowData $rowResult -LabelName "Diskspace Utilization Monitor on (?<ServerName>.+) BrowsableValue2"

    $dp_3 = Get-ParsedGroupData -RowData $rowResult -LabelName "Diskspace Utilization Monitor on (?<ServerName>.+) BrowsableValue3"

    $dp_4 = Get-ParsedGroupData -RowData $rowResult -LabelName "Diskspace Utilization Monitor on (?<ServerName>.+) BrowsableValue4"

    $dp_counters = Get-ParsedGroupData -RowData $rowResult -LabelName "Diskspace Utilization Monitor on (?<ServerName>.+) Counters In Error"

    #export the data to an excel sheet
    $outputFilePath = "ParsedSiteScopeReport_" + (Get-Date -Format "MM_dd_yyyy_hh_mm_ss") + ".xlsx"

    $uptimeSummary | Export-Excel -Path $outputFilePath -WorksheetName "Uptime Summary" -TableName "UptimeSummaryTable" -AutoSize
    $measurementSummary | Export-Excel -Path $outputFilePath -WorksheetName "Measurement Summary" -TableName "MeasureSummaryTable" -AutoSize
    $errorTimeSummary | Export-Excel -Path $outputFilePath -WorksheetName "errorTimeSummary" -TableName "ErrorTimeSummaryTable" -AutoSize
    $cpuResult | Export-Excel -Path $outputFilePath -WorksheetName "CPU" -TableName "CPUTable" -AutoSize
    $pm_virtual_mem_or_swap_space_used_per | Export-Excel -Path $outputFilePath -WorksheetName "PM_VM or Swap Space Used %" -TableName "PMVMorSwapPerTable" -AutoSize
    $pm_virtual_mem_or_swap_space_mb_free | Export-Excel -Path $outputFilePath -WorksheetName "PM_VM or Swap Space MB Free" -TableName "PMVMorSwapMBTable" -AutoSize
    $pm_mb_free | Export-Excel -Path $outputFilePath -WorksheetName "PM_MB Free" -TableName "PMMBFreeTable" -AutoSize
    $pm_per_used | Export-Excel -Path $outputFilePath -WorksheetName "PM_Percent Used" -TableName "PMPerUsedTable" -AutoSize
    $pm_pm_mb_free | Export-Excel -Path $outputFilePath -WorksheetName "PM_Physical Memory MB Free" -TableName "PMPMMemMBTable" -AutoSize
    $pm_pm_used_per | Export-Excel -Path $outputFilePath -WorksheetName "PM_Physical Memory Used %" -TableName "PMPMUserPerTable" -AutoSize
    $pm_pages_sec | Export-Excel -Path $outputFilePath -WorksheetName "PM_ Pages per Sec" -TableName "PMPagesSecTable" -AutoSize
    $pingDetails | Export-Excel -Path $outputFilePath -WorksheetName "Ping Monitor Round Trip Time" -TableName "PingDetailsTable" -AutoSize
    $vm_virtual_mem_or_swap_space_used_per | Export-Excel -Path $outputFilePath -WorksheetName "VM_VM or Swap Space Used %" -TableName "VMVMSwapSpacePerTable" -AutoSize
    $vm_virtual_mem_or_swap_space_mb_free | Export-Excel -Path $outputFilePath -WorksheetName "VM_VM or Swap Space MB Free" -TableName "VMVMSwapSpaceMBFreeTable" -AutoSize
    $vm_mb_free | Export-Excel -Path $outputFilePath -WorksheetName "VM_MB Free" -TableName "VMMBFreeTable" -AutoSize
    $vm_per_used | Export-Excel -Path $outputFilePath -WorksheetName "VM_Percent Used" -TableName "VMPerUsedTable" -AutoSize
    $vm_pm_mb_free | Export-Excel -Path $outputFilePath -WorksheetName "VM_Physical Memory MB Free" -TableName "VMPMMBFreeTable" -AutoSize
    $vm_pm_used_per | Export-Excel -Path $outputFilePath -WorksheetName "VM_Physical Memory Used %" -TableName "VMPMUsedPerTable" -AutoSize
    $vm_pages_sec | Export-Excel -Path $outputFilePath -WorksheetName "VM_ Pages per Sec" -TableName "VMPagesSecTable" -AutoSize
    $dp_1 | Export-Excel -Path $outputFilePath -WorksheetName "DiskSpace Util Browsable1" -TableName "DiskSpaceBrowse1Table" -AutoSize
    $dp_2 | Export-Excel -Path $outputFilePath -WorksheetName "DiskSpace Util Browsable2" -TableName "DiskSpaceBrowse2Table" -AutoSize
    $dp_3 | Export-Excel -Path $outputFilePath -WorksheetName "DiskSpace Util Browsable3" -TableName "DiskSpaceBrowse3Table" -AutoSize
    $dp_4 | Export-Excel -Path $outputFilePath -WorksheetName "DiskSpace Util Browsable4" -TableName "DiskSpaceBrowse4Table" -AutoSize
    $dp_counters | Export-Excel -Path $outputFilePath -WorksheetName "DiskSpace Util Counter In Error" -TableName "DiskSpaceCounterTable" -AutoSize

}


function Get-ParsedUptimeSummary {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object[]]$SummaryRow
    )

    $result = @()

    for ($i = 0; $i -lt $SummaryRow.Count; $i++) {
        $obj = New-Object -TypeName psobject

        $rowProperties = $SummaryRow[$i] | Get-Member -MemberType Property | Select-Object -ExpandProperty Name

        foreach ($property in $rowProperties) {
            $obj | Add-Member -Name "$property" -Value ($SummaryRow[$i].$property) -MemberType NoteProperty
        }

        $result += $obj
    }

    return $result |
    Select-Object @{Label="Name";Expression={$_.name}}, @{Label="Uptime %";Expression={$_.uptime}}, @{Label="Error %";Expression={$_.error}}, @{Label="Warning %";Expression={$_.warning}}, @{Label="Last";Expression={$_.last}}
}


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


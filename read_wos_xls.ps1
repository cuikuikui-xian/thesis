$sourcePath = 'C:\Users\32933\Downloads\savedrecs.xls'
$outputPath = Join-Path $PSScriptRoot 'savedrecs_extracted.json'

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($sourcePath, 0, $true)

    $sheets = @()
    foreach ($sheet in $workbook.Worksheets) {
        $used = $sheet.UsedRange
        $rows = $used.Rows.Count
        $cols = $used.Columns.Count
        $matrix = @()
        for ($r = 1; $r -le $rows; $r++) {
            $row = @()
            for ($c = 1; $c -le $cols; $c++) {
                $row += $used.Cells.Item($r, $c).Text
            }
            $matrix += ,$row
        }
        $sheets += [pscustomobject]@{
            name = $sheet.Name
            rows = $rows
            columns = $cols
            values = $matrix
        }
    }

    $sheets | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $outputPath -Encoding UTF8
    Write-Output $outputPath
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $workbook) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
    if ($null -ne $excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

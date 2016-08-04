$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Add()
$workSheet = $workbook.Sheets(1)

for($i=1; $i -le 100; $i++) {
        
    for($j=1; $j -le 100; $j++) {
        $workSheet.Cells.Item($i, $j).Value = "hoge"
        $workSheet.Cells.Item($i, $j).Font.ColorIndex = 3
        $workSheet.Cells.Item($i, $j).Font.Size = 15
    }

}

$workbook.SaveAs("$PSScriptRoot\hoge.xlsx")
$excel.Quit()

$excel = $null
$workbook = $null
$workSheet = $null
[GC]::Collect()
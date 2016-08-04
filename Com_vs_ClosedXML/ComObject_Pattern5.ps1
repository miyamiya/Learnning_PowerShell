$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Add()
$workSheet = $workbook.Sheets(1)

[object[,]]$all = New-Object "System.object[,]" 100,100

for($i=0; $i -lt 100; $i++) {
        
    for($j=0; $j -lt 100; $j++) {
        $all[$i,$j] = "hoge" 
    }

}

$workSheet.Range($workSheet.Cells(1, 1),$workSheet.Cells(100, 100)).Value = $all
$workbook.SaveAs("$PSScriptRoot\hoge.xlsx")
$excel.Quit()
$excel = $null
$workbook = $null
$workSheet = $null
[GC]::Collect()
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Add()
$workSheet = $workbook.Sheets(1)

[object[]]$line = New-Object System.object[] 100

for($i=1; $i -le 100; $i++) {
        
    for($j=0; $j -lt 100; $j++) {
        $line[$j] = "hoge" 
    }

    $workSheet.Range($workSheet.Cells($i, 1), $workSheet.Cells($i, 100)).Value = $line
    $workSheet.Range($workSheet.Cells($i, 1), $workSheet.Cells($i, 100)).Font.ColorIndex = 3
    $workSheet.Range($workSheet.Cells($i, 1), $workSheet.Cells($i, 100)).Font.Size = 15
    $line.Clear();
}

$workbook.SaveAs("$PSScriptRoot\hoge.xlsx")
$excel.Quit()

$excel = $null
$workbook = $null
$workSheet = $null
[GC]::Collect()
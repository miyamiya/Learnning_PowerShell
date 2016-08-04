$ErrorActionPreference = "Stop"
[Reflection.Assembly]::LoadFile("$PSScriptRoot\ClosedXML.dll")
[Reflection.Assembly]::LoadFile("C:\Program Files (x86)\Open XML SDK\V2.0\lib\DocumentFormat.OpenXml.dll")

$workBook = new-object ClosedXML.Excel.XLWorkbook
$workSheet = $workBook.Worksheets.Add("Sheet1")   

[object[]]$line = New-Object System.object[] 100

for($i=1; $i -le 100; $i++) {
        
    for($j=0; $j -lt 100; $j++) {
        $line[$j] = "hoge"
    }

    $workSheet.Cell($i, 1).InsertData(@(,$line)) > $null
    $line.Clear();
}

$workBook.SaveAs("$PSScriptRoot\hoge.xlsx")
$workBook.Dispose()
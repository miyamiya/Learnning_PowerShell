$ErrorActionPreference = "Stop"
[Reflection.Assembly]::LoadFile("$PSScriptRoot\ClosedXML.dll")
[Reflection.Assembly]::LoadFile("C:\Program Files (x86)\Open XML SDK\V2.0\lib\DocumentFormat.OpenXml.dll")

$workBook = new-object ClosedXML.Excel.XLWorkbook
$workSheet = $workBook.Worksheets.Add("Sheet1")   

for($i=1; $i -le 100; $i++) {
        
    for($j=1; $j -le 100; $j++) {
        $worksheet.Cell($i, $j).Value = "hoge"
    }

}

$workBook.SaveAs("$PSScriptRoot\hoge.xlsx")
$workBook.Dispose()
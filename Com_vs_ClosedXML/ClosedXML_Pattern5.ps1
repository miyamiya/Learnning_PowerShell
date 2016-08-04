$ErrorActionPreference = "Stop"
[Reflection.Assembly]::LoadFile("$PSScriptRoot\ClosedXML.dll")
[Reflection.Assembly]::LoadFile("C:\Program Files (x86)\Open XML SDK\V2.0\lib\DocumentFormat.OpenXml.dll")

$workBook = new-object ClosedXML.Excel.XLWorkbook
$workSheet = $workBook.Worksheets.Add("Sheet1")   

[object[]]$all = New-Object System.object[] 100
[string[]]$line = New-Object System.String[] 100

for($i=1; $i -le 100; $i++) {
        
    for($j=0; $j -lt 100; $j++) {
        $line[$j] = "hoge"
    }

    $all[$i-1] = @($line)
    $line.Clear();
}

$workSheet.Cell(1, 1).InsertData($all) > $null
$workBook.SaveAs("$PSScriptRoot\hoge.xlsx")
$workBook.Dispose()
Dim excelFilePath,worksheetName,headerRange
Set Arg = WScript.Arguments

excelFilePath = Arg(0)
worksheetName = Arg(1)
headerRange = Arg(2)

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True 

Set objWorkbook = objExcel.Workbooks.Open(excelFilePath)

Set mainWorksheet = objWorkbook.Sheets(worksheetName)

mainWorksheet.Activate

Set headers = mainWorksheet.Range(headerRange)

For Each headerCell In headers
    headerCell.Font.Bold = true
    headerCell.Interior.Color = RGB(0,255,0)
Next

objWorkbook.Save

objWorkbook.Close 

objExcel.Quit
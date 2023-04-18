Dim xlApp As Microsoft.Office.Interop.Excel.Application
Dim wb As Microsoft.Office.Interop.Excel.Workbook
Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet

Dim strOne As String

' instantiate Excel Application
xlApp = New Microsoft.Office.Interop.Excel.Application

' open excel (Excel file  whole path)

' @params in_fileName String in
wb = xlApp.Workbooks.Open(in_fileName)
xlApp.DisplayAlerts = False

' TryCast(wb.Sheets(in_sheetname), Microsoft.Office.Interop.Excel.Worksheet).Delete()
' get excel sheet instance

' @params in_sheetName String in
xlSheet = TryCast(wb.Sheets(in_sheetName), Microsoft.Office.Interop.Excel.Worksheet)
'This will give you the value of cell A2 in io_strOne
' @params io_strOne String in/out
io_strOne = xlSheet.Range("A2").Value.toString()  

' https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
xlApp.DisplayAlerts = True

'https://learn.microsoft.com/en-us/office/vba/api/excel.application.visible
xlApp.Visible = False

' save excel
wb.Save()
' close excel connect
wb.Close()

MyFile = Wscript.Arguments.Item(0)
WScript.Echo MyFile
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open(MyFile)
Set objWorksheet = objWorkbook.Worksheets(1)
'objWorksheet.Cells(1, 1).Value = Now
objWorkbook.Save()
objExcel.Quit

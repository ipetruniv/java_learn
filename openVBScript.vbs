MyFile = Wscript.Arguments.Item(0)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.DisplayAlerts = False
'Set objWorkbook = objExcel.Workbooks.Open("C:\\Users\\Yevhen_Veklyn\\Desktop\\zz\\New folder\\Digital_R2_MSA_Import_MD_4261.xlsm", 0, False)
Set objWorkbook = objExcel.Workbooks.Open(MyFile, 0, False)
CreateNoWindow = true
objExcel.Run "ThisWorkbook.SetupProductSheets"


'objExcel.SendKeys "{ENTER}"

'Set cb = objExcel.Shapes("chkOnlyChangedValidation")
'If cb.OLEFormat.Object.Value = 1 Then
'      ' MsgBox "Checkbox is Checked"
' Else
'        MsgBox "Checkbox is not Checked"
'   End If
'cmdOK_Click

objWorkbook.Save
objExcel.ActiveWorkbook.Close

objExcel.Application.Quit
WScript.Quit

'Set objExcel = CreateObject("Excel.Application")
'objExcel.Application.Run "'C:\Users\Ryan\Desktop\Sales.xlsm'!SalesModule.SalesTotal"
'objExcel.DisplayAlerts = False
'objExcel.Application.Quit
'Set objExcel = Nothing

Set objExcel = CreateObject("Excel.Application")

If WScript.Arguments.Count > 0 Then
	filePath = WScript.Arguments.Item(0)
	cities = "NULL"
	On Error Resume Next 
	Set objWorkbook = objExcel.Workbooks.Open(filePath)

	If Err.Number <> 0 Then 
		Msgbox Err.Description 
		Err.Clear
	Else
		'objExcel.Application.Visible = True
		objWorkbook.WorkSheets(1).Activate
		cities = ""
		For i=1 to objExcel.ActiveWorkbook.Sheets.Count
		   cities = objWorkbook.WorkSheets(i).Name + "     " + cities
		Next
		
		objExcel.ActiveWorkbook.Close	
	End If
Else
	MsgBox "Please pass a parameter to this script"
End if

objExcel.Application.Quit
WScript.StdOut.WriteLine(cities)
WScript.Quit
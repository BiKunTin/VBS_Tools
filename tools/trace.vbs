dim objFso: set objFso = CreateObject("Scripting.FileSystemObject")
dim path,ws,wb
strPath = objFso.GetAbsolutePathName(".")
strPath = objFso.GetParentFolderName(strPath)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objFolder = objFso.GetFolder (strPath +"\output\trace\")
For Each objFile In objFolder.Files
If objFso.GetExtensionName (objFile.Path) = "xlsx" Then
Set objWorkbook = objExcel.Workbooks.Open(objFile.Path)
	'For Each Sht In objWorkbook.Worksheets
Set wb = objExcel.Workbooks.Open(strPath+"\output\report_trace.xlsx")
objWorkbook.Worksheets(3).Copy wb.Sheets(1)
Dim k
k = wb.Sheets.Count
For i = k To 3 Step -1
    wb.Sheets(i).Delete
Next
wb.Save
wb.Close
objWorkbook.Close True
End If
Next
objExcel.Quit
MsgBox "Trace Report Finish"
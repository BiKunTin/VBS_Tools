dim fso: set fso = CreateObject("Scripting.FileSystemObject")
dim path,ws,wb
strPath = fso.GetAbsolutePathName(".")
strPath = fso.GetParentFolderName(strPath)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objFolder = fso.GetFolder (strPath +"\output\testcase\")
For Each objFile In objFolder.Files
If fso.GetExtensionName (objFile.Path) = "xlsx" Then
Set objWorkbook = objExcel.Workbooks.Open(objFile.Path)
Set wb = objExcel.Workbooks.Open(strPath+"\output\report_testcase.xlsx")
objWorkbook.Worksheets(2).Copy wb.Sheets(1)
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
MsgBox "Testcase Report Finish"
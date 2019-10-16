Set objExcel = CreateObject("Excel.Application")
dim fso: set fso = CreateObject("Scripting.FileSystemObject")
dim path,ws
path = fso.GetAbsolutePathName(".") 
path = path + "\tools\"
objExcel.Workbooks.Open(path + "MIB3.xlsm")
objExcel.Application.Visible = False
objExcel.Application.Run "Make_Report"
objExcel.DisplayAlerts = False
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")
objShell.Run Chr(34) & path &"trace.vbs" & Chr(34)
objExcel.Application.Quit
MsgBox "Report Finish"
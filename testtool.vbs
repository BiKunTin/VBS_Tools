dim fso: set fso = CreateObject("Scripting.FileSystemObject")
Dim objShell
Dim path
path = fso.GetAbsolutePathName(".") 
Set objShell = Wscript.CreateObject("WScript.Shell")
objShell.Run path & "\trace.vbs" 
' Using Set is mandatory
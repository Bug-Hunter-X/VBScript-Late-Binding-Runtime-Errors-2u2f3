Early Binding and Error Handling:
```vbscript
On Error GoTo ErrorHandler

Dim objExcel As Object
Set objExcel = CreateObject("Excel.Application")

MsgBox objExcel.Version

Exit Sub

ErrorHandler:
MsgBox "Error: " & Err.Description
End Sub
```
Alternatively, using early binding (requires adding a reference to the Excel library in the VBScript project):
```vbscript
Dim objExcel As Excel.Application
Set objExcel = CreateObject("Excel.Application")

MsgBox objExcel.Version
```
Late Binding: VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where version mismatches or incorrect references can occur.  The script might not throw an error during development but fail unpredictably in a different environment.
```vbscript
Dim objExcel
Set objExcel = CreateObject("Excel.Application")
'Error if Excel is not installed or the version is incompatible
MsgBox objExcel.Version
```
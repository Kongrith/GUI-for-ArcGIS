Dim wshShell, fso, loc, cmd

Set fso = CreateObject("Scripting.FileSystemObject")
loc = fso.GetAbsolutePathName(".")
WScript.Echo loc

'~ cmd = "%ComSpec% /k C:\Languages\Python\python.exe " + loc + "\test.py"
cmd = "C:\Languages\Python\python.exe " + loc + "\test.py"
WScript.Echo cmd

Set wshShell = CreateObject("WScript.Shell")
wshShell.Run cmd


------------------- (เรียก excel และสั่ง cursur ไปที่ sub )-----------------------------------

Function OpenLink ( [Substation_point.hyperlink]  , [search.Index]  )
Dim objExcel, objSheet, objRange, path, a
path = [Substation_point.hyperlink] 
a = [search.Index] 

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open( path)
Set objSheet = objWorkbook.Sheets("Con Grid")
objexcel.Visible=TRUE

Set objRange = objExcel.Range(a)
objRange.Activate


End Function
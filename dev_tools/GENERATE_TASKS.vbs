'GENERATE_TASKS.vbs

Option Explicit

'Use security settings of Excel application
Const msoAutomationSecurityForceDisable  = 1

Dim app, autoSec, doc, fileSys
Set fileSys = CreateObject("Scripting.FileSystemObject")
Set app = CreateObject("Excel.Application")

app.DisplayAlerts = False
autoSec = app.AutomationSecurity
app.AutomationSecurity = msoAutomationSecurityForceDisable
Err.Clear
Set doc = app.Workbooks.Open(WScript.Arguments(0), False, True)

If Err = 0 Then
    Call app.Run(doc.name + "!" + "API_AddNewTask", "name", 1, 2)
	doc.Close True
Else
    WScript.Echo("Could not open workbook '" & WScript.Arguments(0) & "'")
End If

app.AutomationSecurity = autoSec
app.Quit
WScript.Quit Err
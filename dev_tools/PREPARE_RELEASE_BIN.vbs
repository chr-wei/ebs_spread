'PREPARE_RELEASE_BIN.vbs
'Input args
'   Arguments(0):   Path of EbsSpread excel file to be prepared for release
'   Arguments(1):   Path of destination file
Option Explicit

Dim xlApp, storedSec, doc, fileSys
Set fileSys = CreateObject("Scripting.FileSystemObject")
Set xlApp = CreateObject("Excel.Application")

Dim wbInPath, wbOutPath, releaseVer
wbInPath = WScript.Arguments(0)
wbOutPath = WScript.Arguments(1)

xlApp.DisplayAlerts = False
storedSec = xlApp.AutomationSecurity

'Use security settings of Excel application
xlApp.AutomationSecurity = 2

On error resume next
Err.Clear
'Open spreadsheet, do not update links, open read only
Set doc = xlApp.Workbooks.Open(wbInPath, False, True)

If Err = 0 Then
    'Add tasks with args:
    '                                                 name,        comment,               tag1,        tag2,      kanban list,  task estimate,  total time, due date
    Call xlApp.Run(doc.name + "!" + "API_AddNewTask", "Buy pizza", "Most delicious food", "fast food", "Italian", "To do",      0.5,            1.25,       Date() + 10)
	
    If Err = 0 then
        Call doc.Close(True, wbOutPath)
    else 
        WScript.Echo("Could not save workbook '" & wbOutPath & "'")
    End if
Else
    WScript.Echo("Could not open workbook '" & wbInPath & "'")
End If

xlApp.DisplayAlerts = True
xlApp.AutomationSecurity = storedSec
xlApp.Quit
WScript.Quit Err
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

'Comment to see more errors
'On error resume next
Err.Clear
'Open spreadsheet, do not update links, open read only
Set doc = xlApp.Workbooks.Open(wbInPath, False, True)

If Err <> 0 Then
    WScript.Echo("Could not open workbook '" & wbInPath & "'")
else 
    'Run API calls

    'Delete existing data
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_DeleteAllTasks")
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_DeleteAllEbsSheets")
    
    'Add tasks with args
    '                                                              name,                                comment,                    tag1,           tag2,           kanban list, task estimate, total time,     due date
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_AddNewTask", "Buy strawberries",                 "",                         "fruits",       "",             "To do",            6.00,        0.00,      cDate(Date() + 0))
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_AddNewTask", "Take a milk carton",               "Watch best before note",   "fresh",        "",             "To do",            0.25,        0.00,      cDate(Date() + 3))
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_AddNewTask", "Buy rice",                         "",                         "",             "",             "Done",             2.00,        1.00,      cDate(Date() + 2))
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_AddNewTask", "Find shopping cart",               "",                         "",             "",             "Done",             4.00,        7.75,      cDate(0))
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_AddNewTask", "Have a chat with the store owner", "",                         "",             "important",    "Done",             1.00,        0.50,      cDate(Date() + 0))
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_AddNewTask", "Buy soy beans",                    "",                         "fast food",    "",             "In progress",      0.50,        1.25,      cDate(Date() + 0))
    Call xlApp.Run(doc.name + "!" + "API_Functions.API_AddNewTask", "Buy apples",                       "",                         "fruits",       "important",    "Testing/Review",   3.00,        0.25,      cDate(Date() + 0))

    If Err <> 0 then
        WScript.Echo("One or more API call(s) failed. Remove 'On error resume next' statement to debug'")
    else
        Call doc.Close(True, wbOutPath)
        If Err <> 0 then
            WScript.Echo("Could not save workbook '" & wbOutPath & "'")
        else
            WScript.Echo("Release was successfully exported to '" & wbOutPath & "'")
        end if
    End if
End If

xlApp.DisplayAlerts = True
xlApp.AutomationSecurity = storedSec
xlApp.Quit
Set xlApp = Nothing
WScript.Quit Err
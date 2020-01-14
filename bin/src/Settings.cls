VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub UI_StoreSelectedCalendarId()
    'This function stores information of a selected outlook calendar in the settings sheet.
    'The stored id can be copied to a contributors settings row to retrieve calendar information.
    
    Dim sheet As Worksheet
    Set sheet = SettingUtils.GetSettingsSheet
    
    If sheet Is Nothing Then Exit Sub
    
    Dim calId As String
    Dim storId As String
    Dim calPath As String
    calId = CalendarUtils.GetSelectedCalendarId(storId, calPath)
    
    If StrComp(calId, "") <> 0 Then
        Dim idCell As Range
        Call Utils.GetSingleDataCellVal(sheet, Constants.CAL_ID_TEMP_HEADER, idCell)
        If Not idCell Is Nothing Then idCell.Value = calId
    End If
    
    If StrComp(storId, "") <> 0 Then
        Dim storCell As Range
        Call Utils.GetSingleDataCellVal(sheet, Constants.STORE_ID_TEMP_HEADER, storCell)
        If Not storCell Is Nothing Then storCell.Value = storId
    End If
    
    If StrComp(calPath, "") <> 0 Then
        Dim pathCell As Range
        Call Utils.GetSingleDataCellVal(sheet, Constants.CAL_PATH_TEMP_HEADER, pathCell)
        If Not pathCell Is Nothing Then pathCell.Value = calPath
    End If
End Sub



Sub UI_EnableEvents()
    'Enable all events in case they got disabled before. Can happen when errors ocur and function calls are not finished
    Application.EnableEvents = True
    Application.screenUpdating = True
End Sub



Sub UI_ExportTasks()
    'Export all visible tasks of planning sheet list. Data will be exported to a new workbook
    Call Utils.RunTryCatchedCall(f:="ExportImportUtils.ExportVisibleTasks", screenUpdating:=False)
End Sub



Sub UI_ImportTasks()
    'Import all tasks of a workbook which name is specified in the setting sheet.
    
    'Read the workbook's name and check whether it is opened
    Dim importWorkbookName As String
    importWorkbookName = Utils.GetSingleDataCellVal(SettingUtils.GetSettingsSheet, Constants.IMPORT_SHEET_HEADER)

    Dim sourceWb As Workbook
    On Error Resume Next
    Set sourceWb = Workbooks(importWorkbookName)
    
    If Err.Number <> 0 Then
        'Workbook is not opened
        Debug.Print "Importing from workbook '" & importWorkbookName & "' failed."
        Err.Clear
        Exit Sub
    Else
        On Error GoTo 0
    End If
    
    'Workbook available. Import tasks' data
    Call Utils.RunTryCatchedCall(f:="ExportImportUtils.ImportTasks", arg1:=sourceWb, screenUpdating:=False)
End Sub

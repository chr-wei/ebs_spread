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
    Dim calPath As String
    calId = CalendarUtils.GetSelectedCalendarId(calPath)
    
    If StrComp(calId, "") <> 0 Then
        Dim idCell As Range
        
        Call Utils.GetSingleDataCellVal(sheet, Constants.CAL_ID_HEADER, idCell)
        If Not idCell Is Nothing Then idCell.Value = calId
    End If
    
    If StrComp(calPath, "") <> 0 Then
        Dim pathCell As Range
        Call Utils.GetSingleDataCellVal(sheet, Constants.CAL_PATH_HEADER, pathCell)
        If Not idCell Is Nothing Then pathCell.Value = calPath
    End If
End Sub



Sub UI_EnableEvents()
    'Enable all events in case they got disabled before. Can happen when errors ocur and function calls are not finished
    Application.EnableEvents = True
    Application.screenUpdating = True
End Sub


Attribute VB_Name = "SettingUtils"
'  This macro collection lets you organize your tasks and schedules
'  for you with the evidence based schedule (EBS) approach by Joel Spolsky.
'
'  Copyright (C) 2020  Christian Weihsbach
'  This program is free software; you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation; either version 3 of the License, or
'  (at your option) any later version.
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'  You should have received a copy of the GNU General Public License
'  along with this program; if not, write to the Free Software Foundation,
'  Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301  USA
'
'  Christian Weihsbach, weihsbach.c@gmail.com

Option Explicit

Enum WorkingTime
    ceWorkingStart = -1
    ceWorkingEnd = 1
End Enum

Enum ApptOnOffset
    ceOnset = -1
    ceOffset = 1
End Enum

Enum ColorType
    ceAccentColor = 0
    ceCommonColor = 1
    ceLightColor = 2
End Enum



Function GetSettingsSheet() As Worksheet
    'Return the setting sheet (fixed name)
    '
    'Output args:
    '  GetSettingSheet:    Handle to the setting sheet
    
    Set GetSettingsSheet = ThisWorkbook.Worksheets(Constants.SETTING_SHEET_NAME)
End Function



Function GetContributorSettingColumn(colIdentifier As Variant, rowIdentifier As ListRowSelect) As Range
    'Wrapper to read a specified column of contributor settings
    '
    'Input args:
    '  colIdentifier:                  An identifier specifying the column one whishes to extract. Can be a cell inside the column or the header string
    '  rowIdentifier:                  An identifier specifying whether to return the whole list, only the header or only the data range (without headers)
    '
    'Output args:
    '  GetContributorSettingColumn:    Range of the selected cells of the column
    
    Set GetContributorSettingColumn = Utils.GetListColumn(SettingUtils.GetSettingsSheet, Constants.CONTRIBUTOR_LIST_NAME, colIdentifier, rowIdentifier)
End Function



Function IntersectContributorAndSettingColumn(contributor As String, colIdentifier As String) As Range
    'Wrapper to read a specific setting for a contributor.
    '
    'Input args:
    '  contributor:    The name of the contributor you want to read the setting for
    '  colIdentifier:  The setting / column you want to read
    '
    'Output args:
    'IntersectContributorAndSettingColumn: The cell containing the setting for a contributor
    
    'Init output
    Set IntersectContributorAndSettingColumn = Nothing
    
    'Check args
    If StrComp(contributor, "") = 0 Or StrComp(colIdentifier, "") = 0 Then Exit Function
    
    'Get a cell of a data column to a specific contributor
    Dim contributorCell As Range
    If SettingUtils.GetContributorCell(contributor, contributorCell) Then
        Set IntersectContributorAndSettingColumn = _
            Utils.IntersectListColAndCells(SettingUtils.GetSettingsSheet, Constants.CONTRIBUTOR_LIST_NAME, colIdentifier, contributorCell)
    End If
End Function



Function GetContributorCell(contributor As String, Optional ByRef contributorCell) As Boolean
    'Returns true if contributor entry can be found in list. Additionally the cell is returned if a byref arg is given.
    '
    'Input args:
    '  contributor: The contributor name
    'Output args:
    '  contributorCell: The cell containing the contributor name
    '  GetContributorCell: True if the contributor can be found in the list
    
    'Init output
    GetContributorCell = False
    
    'Check args
    If StrComp(contributor, "") = 0 Then Exit Function
    
    Dim data As Range
    Set data = SettingUtils.GetContributorSettingColumn(Constants.CONTRIBUTOR_HEADER, ceData)
    Set contributorCell = Base.FindAll(data, contributor)
    
    'Set the output / getting cell successful?
    GetContributorCell = Not contributorCell Is Nothing
End Function



Function GetContribWorkTimeSetting(contributor As String, timeSelector As WorkingTime) As Date
    Const FN As String = "GetContribWorkTimeSetting"
    'Read the contributor's work time. Fallback constants are used if the contributor can not be found in the list.
    '
    'Input args:
    '  contributor:            The contributor
    '  timeSelector:           Selector specifying the working time one wants to read (start or end)
    '
    'Output args:
    '  GetContribWorkTimeSetting: The working time a contributor starts or ends working
    '
    'Init output
    GetContribWorkTimeSetting = CDate(0)
    
    Dim contributorWorkingTime As Date
    Dim hourCell As Range
    
    Select Case timeSelector
        Case WorkingTime.ceWorkingStart
            Set hourCell = IntersectContributorAndSettingColumn(contributor, Constants.WORKING_HOURS_START_HEADER)
        Case WorkingTime.ceWorkingEnd
            Set hourCell = IntersectContributorAndSettingColumn(contributor, Constants.WORKING_HOURS_END_HEADER)
    End Select
    
    Dim fallback As Boolean
    fallback = False
    
    If hourCell Is Nothing Then
        'Fallback to constants if contributor is not in the list of settings
        fallback = True
    Else
        If StrComp(hourCell.Value, "") = 0 Then
            'Fallback to constants if setting is empty
            fallback = True
        Else
            fallback = False
        End If
    End If
    
    If fallback Then
        Select Case timeSelector
            Case WorkingTime.ceWorkingStart
                contributorWorkingTime = Constants.WORKING_HOURS_START
                Call MessageUtils.HandleMessage("Falling back to standard working hour start [" & _
                    Constants.WORKING_HOURS_START & "] for contributor" & contributor & "'", _
                        ceInfo, FN)
            Case WorkingTime.ceWorkingEnd
                contributorWorkingTime = Constants.WORKING_HOURS_END
                Call MessageUtils.HandleMessage("Falling back to standard working hour end [" & _
                    Constants.WORKING_HOURS_END & "] for contributor" & contributor & "'", _
                        ceInfo, FN)
        End Select
    Else
        'Read the setting from the cell
        contributorWorkingTime = CDate(hourCell.Value)
    End If
    
    GetContribWorkTimeSetting = contributorWorkingTime
End Function



Function GetContributorMailSetting(contributor As String) As String
    'Get the mail address of a contributor
    '
    'Input args:
    '  contributor: The name of the contributor
    'Init output
    GetContributorMailSetting = ""
    
    'Check args
    
    If StrComp(contributor, "") = 0 Then Exit Function
    
    Dim mailCell As Range
    
    Set mailCell = SettingUtils.IntersectContributorAndSettingColumn(contributor, MAIL_HEADER)
    
    If mailCell Is Nothing Then
        Exit Function
    Else
        GetContributorMailSetting = mailCell.Value
    End If
End Function



Function GetImportedTaskPostfixSetting() As String
    'Get the postfix setting string that is added to imported tasks
    '

    'Init output
    GetImportedTaskPostfixSetting = ""
    
    Dim postfixCell As Range
    Dim postfix As String
    postfix = Utils.GetSingleDataCellVal(SettingUtils.GetSettingsSheet, Constants.IMPORTED_TASK_POSTFIX_HEADER, postfixCell)
    
    If postfixCell Is Nothing Then
        Exit Function
    Else
        GetImportedTaskPostfixSetting = postfix
    End If
End Function



Function GetContributorCalIdSetting(contributor As String, Optional ByRef storId As String) As String
    'Get the calendar id of the contributor
    '
    'Input args:
    '  contributor: The name of the contributor
    
    'Init output
    GetContributorCalIdSetting = ""
    storId = ""
    
    'Check args
    If StrComp(contributor, "") = 0 Then Exit Function
    
    Dim idCell As Range
    Dim storeIdCell As Range
    
    Set idCell = SettingUtils.IntersectContributorAndSettingColumn(contributor, Constants.CAL_ID_HEADER)
    Set storeIdCell = SettingUtils.IntersectContributorAndSettingColumn(contributor, Constants.STORE_ID_HEADER)
    
    If Not idCell Is Nothing Then
        GetContributorCalIdSetting = Trim(idCell.Value)
    End If
    
    If Not storeIdCell Is Nothing Then
        storId = Trim(storeIdCell.Value)
    End If
End Function



Function GetContributorGetWorkDoneCat(contributor As String) As String
    'Return a category tag a user can specify in the calendar to mark it as 'get work done'. The tag makes the time calculating algorithm
    'skip the tagged calendar entry and use its time as 'free to work'. A fallback mechanism is used to always retrieve a value even if
    'no specific setting is set.
    '
    'Input args:
    '  contributor: The name of the contributor one wants to receive the getting work done category for
    '
    'Output args:
    '  GetContributorGetWorkDoneCat: The category / categories returned. They have to be in [] brackets [cat] or [cat1, cat2]
    
    'Init output
    GetContributorGetWorkDoneCat = ""
    
    'Check args
    If StrComp(contributor, "") = 0 Then Exit Function
    
    Dim catCell As Range
    
    Set catCell = SettingUtils.IntersectContributorAndSettingColumn(contributor, GETTING_WORK_DONE_CAT_HEADER)
    
    If catCell Is Nothing Then
        'Fallback to standard value
        GetContributorGetWorkDoneCat = Constants.STANDARD_GETTING_WORK_DONE_CATEGORY
    Else
        GetContributorGetWorkDoneCat = catCell.Value
    End If
End Function



Function GetFirstDayOfWeekSetting() As Integer
    'Return first day of week setting. A fallback mechanism is used to always retrieve a value even if
    'no specific setting is set.
    '
    'Output args:
    '  GetFirstDayOfWeekSetting: The weekday as vbSunday to vbMonday
    
    'Init output
    GetFirstDayOfWeekSetting = -1
      
    Dim firstDayOfWeek As Integer
    
    firstDayOfWeek = SettingUtils.GetWeekdayVal(Utils.GetSingleDataCellVal(SettingUtils.GetSettingsSheet, Constants.FIRST_DAY_OF_WEEK_HEADER))

    If firstDayOfWeek = -1 Then
        'Fallback to standard value
        GetFirstDayOfWeekSetting = Constants.FIRST_DAY_OF_WEEK
    Else
        GetFirstDayOfWeekSetting = firstDayOfWeek
    End If
End Function



Function GetContributorApptOnOffset(contributor As String, onOffsetSelector As ApptOnOffset) As Double
    'Returns the appointment on- and offset values. with the on- and offset one can specify how long prior and / or after an appointment
    'time is needed to prepare the meeting / get back to work
    '
    'Input args:
    '  contributor:        The name of the contributor you want to read the setting for
    '  onOffsetSelector:   Selector specifying whether you want to read the onset setting (time to prepare the meeting) or the offset
    '                      (time to get back to work meeting) value
    '
    'Output args:
    '  GetContributorApptOnOffset: Time value in (h)
    
    'Init output
    GetContributorApptOnOffset = -1
    
    'Check args
    If StrComp(contributor, "") = 0 Then Exit Function
    
    Dim cell As Range
    Dim onOffset As Double
    
    Select Case onOffsetSelector
        Case ApptOnOffset.ceOnset
            Set cell = SettingUtils.IntersectContributorAndSettingColumn(contributor, APPT_ONSET_HEADER)
        Case ApptOnOffset.ceOffset
            Set cell = IntersectContributorAndSettingColumn(contributor, APPT_OFFSET_HEADER)
    End Select
    
    Dim fallback As Boolean
    fallback = False
    
    If cell Is Nothing Then
        fallback = True
    Else
        If StrComp(cell.Value, "") = 0 Or Not IsNumeric(cell.Value) Then
            'No entry found in cell - fallback to constants
            fallback = True
        Else
            fallback = False
        End If
    End If
    
    If fallback Then
        Select Case onOffsetSelector
            Case ApptOnOffset.ceOnset
                onOffset = Constants.APPOINTMENT_ONSET_HOURS
            Case ApptOnOffset.ceOffset
                onOffset = Constants.APPOINTMENT_OFFSET_HOURS
        End Select
    Else
        onOffset = CDbl(cell.Value)
    End If
    
    GetContributorApptOnOffset = onOffset
End Function



Function GetContributorWorkingDays(contributor As String) As Double()
    'Function returns the working days of a contributor. By default days from Monday to Friday are returned.
    '
    'Input arguments:
    '  contributor: Name of the contributor
    '
    'Output arguments:
    '  GetContributorWorkingDays: Array containing the workdays as numbers (starting from sunday (1) to saturday (7)
    
    'Init output
    Dim workingDays() As Double
    GetContributorWorkingDays = workingDays
    
    'Check args
    If StrComp(contributor, "") = 0 Then Exit Function
    
    Dim wdsCell As Range
    
    Set wdsCell = SettingUtils.IntersectContributorAndSettingColumn(contributor, WORKING_DAYS_HEADER)
    
    If wdsCell Is Nothing Then
        GetContributorWorkingDays = Utils.CopyVarArrToDoubleArr(Constants.STANDARD_WORKING_DAYS)
    Else
        Dim wdSerialized As String
        wdSerialized = wdsCell.Value
        
        If StrComp(wdSerialized, "") <> 0 Then
            Dim workingDaysSet() As String
            workingDaysSet = Utils.CopyVarArrToStringArr(Utils.DeserializeArray(wdSerialized))
            ReDim workingDays(0 To UBound(workingDaysSet))
            
            Dim dayIdx As Integer
            'Read the workdays. Serialized array in settings row should look like this: {Mon; Tue; Wed; Thu; Fri}
            
            For dayIdx = 0 To UBound(workingDaysSet)
                workingDays(dayIdx) = SettingUtils.GetWeekdayVal(workingDaysSet(dayIdx))
            Next dayIdx
            GetContributorWorkingDays = workingDays
        End If
    End If
End Function



Function GetWeekdayVal(id As String) As Integer
    'This function returns vb constant values for the following three character day ids:
    'Sun, Mon, Tue, Wed, Thu, Fri, Sat'
    Select Case Trim(id)
        Case "Sun"
            GetWeekdayVal = vbSunday
        Case "Mon"
            GetWeekdayVal = vbMonday
        Case "Tue"
            GetWeekdayVal = vbTuesday
        Case "Wed"
            GetWeekdayVal = vbWednesday
        Case "Thu"
            GetWeekdayVal = vbThursday
        Case "Fri"
            GetWeekdayVal = vbFriday
        Case "Sat"
            GetWeekdayVal = vbSaturday
        Case Else
            'Invalid day
            GetWeekdayVal = -1
    End Select
End Function



Function GetColors(colorT As ColorType) As Long
    'Get colors from the setting sheet (colored cell) with cell identifier left of it.
    '
    'Input args:
    '  colorT: Type of the xolor one wants to read. Sets the cell identifier to get the correct color
    '
    'Output args:
    '  GetColors:  The color you requested
    
    GetColors = 0
    Dim co As String
    Select Case colorT
        Case ColorType.ceAccentColor
            co = SETTINGS_HIGHLIGHT_COLOR_HEADER
        Case ColorType.ceCommonColor
            co = SETTINGS_COMMON_COLOR_HEADER
        Case ColorType.ceLightColor
            co = SETTINGS_LIGHT_COLOR_HEADER
    End Select
    
    Dim dataCell As Range
    Call Utils.GetSingleDataCellVal(SettingUtils.GetSettingsSheet, co, dataCell)
    
    If dataCell Is Nothing Then
        Exit Function
    Else
        GetColors = CLng(dataCell.Interior.color)
    End If
End Function

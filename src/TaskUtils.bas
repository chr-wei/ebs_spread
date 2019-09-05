Attribute VB_Name = "TaskUtils"
'  This macro collection lets you organize your tasks and schedules
'  for you with the evidence based design (EBS) approach by Joel Spolsky.
'
'  Copyright (C) 2019  Christian Weihsbach
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



Function HandleFollowHyperlink(ByVal Target As hyperlink)
    Dim clickedCell As Range
    Dim sheet As Worksheet
    
    Set clickedCell = Target.parent
    Set sheet = clickedCell.parent
    
    'No hyperlink handling for the task sheet template
    If StrComp(sheet.name, Constants.TASK_SHEET_TEMPLATE_NAME) = 0 Then Exit Function
    
    If Target.subAddress Like Constants.PLANNING_SHEET_NAME + "!*" Then
        'Hide this sheet when clicking the link
        sheet.Visible = XlSheetVisibility.xlSheetHidden
    Else
        Dim accentColor As Long
        accentColor = SettingUtils.GetColors(ceAccentColor)
    
        Select Case clickedCell.Value
            Case Constants.TASK_SHEET_ACTION_ONE_NAME
                Call TaskUtils.SetPlainDelta(clickedCell)
            Case Constants.TASK_SHEET_ACTION_TWOO_NAME
                Call TaskUtils.SetCalendarDelta(clickedCell)
        End Select
        
        'Clear all the accent colors from the cells and highlight the clicked one
        Call TaskUtils.SetClickActionHighlight(sheet, clickedCell)
        
        'Overwrite action of the clicked link and select the clicked cell
        clickedCell.Select
    End If
End Function




Function GetTaskSheet(hash As String) As Worksheet
    'Return the task sheet to a given hash. Has a safety mechanism to always create a task sheet if no sheet could be found.
    If SheetExists(hash) Then
        Set GetTaskSheet = ThisWorkbook.Worksheets(hash)
    Else
        'Add a new worksheet with HASH as name if no sheet exists
        ThisWorkbook.Worksheets(TASK_SHEET_TEMPLATE_NAME).Copy Before:=ThisWorkbook.Worksheets(TASK_SHEET_TEMPLATE_NAME)

        'Go back to overview worksheet as copying jumps to the new sheet
        ThisWorkbook.Worksheets(PLANNING_SHEET_NAME).Activate

        Dim sheet As Worksheet
        Set sheet = ThisWorkbook.Worksheets(TASK_SHEET_TEMPLATE_NAME & " (2)")
        sheet.name = hash

        Call TaskUtils.SetHash(sheet, hash)
        Set GetTaskSheet = sheet
    End If
End Function



Function GetTimeListColumn(sheet As Worksheet, colIdentifier As Variant, rowIdentifier As ListRowSelect) As Range
    'Wrapper to read column of a task list
    Set GetTimeListColumn = Utils.GetListColumn(sheet, TASK_SHEET_TIME_LIST_IDX, colIdentifier, rowIdentifier)
End Function



Function GetEbsListColumn(sheet As Worksheet, colIdentifier As Variant, rowIdentifier As ListRowSelect) As Range
    'Wrapper to read column of a task list
    Set GetEbsListColumn = Utils.GetListColumn(sheet, Constants.TASK_SHEET_EBS_LIST_IDX, colIdentifier, rowIdentifier)
End Function



Function IntersectEbsListColumn(sheet As Worksheet, colIdentifier As Variant, rowIdentifier As Range) As Range
    'Get the intersection of a list column and a cell
    Set IntersectEbsListColumn = Utils.IntersectListColAndCells(sheet, Constants.TASK_SHEET_EBS_LIST_IDX, colIdentifier, rowIdentifier)
End Function



Function IntersectTimeListColumn(sheet As Worksheet, colIdentifier As Variant, rowIdentifier As Range) As Range
    'Wrapper to get the intersection of a list column and a cell
    Set IntersectTimeListColumn = Utils.IntersectListColAndCells(sheet, Constants.TASK_SHEET_TIME_LIST_IDX, colIdentifier, rowIdentifier)
End Function



Function AddNewEbsEntry(sheet As Worksheet, supportPoint As Double, interpolatedEstimate As Double)
    'Adds ebs entry (propability, estimate time) pair to the tasks sheet ebs list
    '
    'Input args:
    '  sheet:                  The task sheet you want to add the estimate for
    '  supportPoint:           The support point (propability) of the estimate
    '  interpolatedEstimate:   The estimate itself
    
    'Check args
    If sheet Is Nothing Then Exit Function
    
    Dim newEntryCell As Range
    
    Dim receivedEntry As Boolean
    Dim newFormattedNumber As String
        
    receivedEntry = GetNewEntry(sheet, Constants.TASK_SHEET_EBS_LIST_IDX, newEntryCell, newFormattedNumber)
    
    If receivedEntry Then
        'If a new entry (with number and cell) could be generated add the data
        Dim supportPointCell As Range
        Dim interpolatedEstCell As Range
        
        Set supportPointCell = TaskUtils.IntersectEbsListColumn(sheet, Constants.SINGLE_SUPPORT_POINT_HEADER, newEntryCell)
        Set interpolatedEstCell = TaskUtils.IntersectEbsListColumn(sheet, Constants.EBS_SELF_TIME_HEADER, newEntryCell)
        
        If Not supportPointCell Is Nothing And Not interpolatedEstCell Is Nothing Then
            newEntryCell.Value = newFormattedNumber
            supportPointCell.Value = supportPoint
            interpolatedEstCell.Value = interpolatedEstimate
        End If
    End If
End Function



Function AddNewTrackingEntry(sheet As Worksheet, Optional durationHours As Double = -1) As Boolean
    'Adds a new line to the time list of the task sheet tracking the elapsed time of tasks. Additionally sets an indicator showing that the entry is
    'currently tracked
    '
    'Input args:
    '  sheet:  The task sheet
    '  durationHours:  A given duration. Immediately finishes the tracking entry and sets the end time of the entry to 'Now'. The start time is set
    '                  to Now (in d) - durationHours / 24
    '
    'Output args:
    '   AddNewTrackingEntry: True if adding entry was successful
    
    'Init output
    AddNewTrackingEntry = False
    
    Dim newEntryCell As Range
    Dim startTimeCol As Range
    Dim startTimeCell As Range
    Dim indicatorCol As Range
    Dim indicatorCell As Range
        
    Dim newFormattedNumber As String
    Dim entryReceived As Boolean
    
    entryReceived = GetNewEntry(sheet, Constants.TASK_SHEET_TIME_LIST_IDX, newEntryCell, newFormattedNumber)
    If entryReceived Then
        Set startTimeCol = TaskUtils.GetTimeListColumn(sheet, Constants.START_TIME_HEADER, ceData)
        Set indicatorCol = TaskUtils.GetTimeListColumn(sheet, Constants.INDICATOR_HEADER, ceData)
        
        Set startTimeCell = Base.IntersectN(newEntryCell.EntireRow, startTimeCol)
        Set indicatorCell = Base.IntersectN(newEntryCell.EntireRow, indicatorCol)

        If indicatorCol Is Nothing Or _
            indicatorCell Is Nothing Or _
            startTimeCell Is Nothing Then
            Exit Function
        End If
        
        'Set the '<current tracker' indicator
        indicatorCol.Value = ""
        indicatorCell.Value = INDICATOR
        
        'Add a new #00000n number to the list
        newEntryCell.Value = newFormattedNumber
        
        Dim startTime As Date
        If durationHours = -1 Then
            'Duration time is invalid. Set start to 'Now' and no end time
            startTime = Now
        Else
            'Substract the hours and convert to days (24 h/day) to add a certain amount of time ending 'Now'. This is only done
            startTime = Now - durationHours / 24
        End If
        
        startTimeCell.Value = startTime
        
        If durationHours <> -1 Then
            'Set end time to 'Now' since start time has been set to a time-backshifted value already
            Call SetEndTimeToSheetTracking(sheet, Now)
        End If
        
        AddNewTrackingEntry = True
    End If
End Function



Function AddXHoursTime(sheet As Worksheet, deltaT As Double)
    'Wrapper to add some hours to the tracking sheet
    '
    'Input args:
    '   sheet:  The task sheet one wants to add the time to
    '   deltaT: The time to add in h
    
    Dim hash As String
    hash = TaskUtils.GetHash(sheet)
    
    If StrComp(hash, "") = 0 Then Exit Function
    
    If PlanningUtils.IsTaskTracking(hash) Then
        'If the selected task is currently tracking, finish it to get a clean state
        PlanningUtils.EndAllTasks
    End If
    
    Call AddNewTrackingEntry(sheet, deltaT)
End Function



Function SetEndTimeToSheetTracking(sheet As Worksheet, Optional time As Date = 0)
    'Ends a current tracking entry by setting 'Now'as end time and reset the indicator '<current'tag
    '
    'Input args:
    '  sheet:  The task sheet
    '  time:   An optional time value to specify a fixed time. If time is 0 then 'Now'will be set
        
    Dim trackingEntry As Range
    Dim endTimeCell As Range
    
    Set trackingEntry = TaskUtils.GetTimeListColumn(sheet, Constants.INDICATOR_HEADER, ceData)
    Set trackingEntry = Base.FindAll(trackingEntry, Constants.INDICATOR)
    
    If trackingEntry Is Nothing Then
        Exit Function
    End If
    
    Set endTimeCell = Utils.IntersectListColAndCells(sheet, Constants.TASK_SHEET_TIME_LIST_IDX, Constants.END_TIME_HEADER, trackingEntry)
    If time = 0 Then
        'If no time value was given take the current time
        endTimeCell.Value = Now
    Else
        'If time value was given take it
        endTimeCell.Value = time
    End If
    
    'Reset the indicator cell
    trackingEntry.Value = ""
    
    'Add special functions to the currently ended tracking entry which let you control delta time calculation (see also 'TaskUtils.HandleFollowHyperlink')
    Call TaskUtils.AddSpecialActionsToTimeListRow(sheet, trackingEntry)
End Function



Function SetTaskName(sheet As Worksheet, name As String)
    'Wrapper to set the task name
    Call Utils.SetSingleDataCell(sheet, TASK_NAME_HEADER, name)
End Function



Function SetEstimate(sheet As Worksheet, userEstimate As Variant)
    'Wrapper to set the user estimate to the task sheet
    'It can be non-numeric in case the user entered an invalid value. The value will still be set to the task sheet for consistency reasons.
    
    Call Utils.SetSingleDataCell(sheet, TASK_ESTIMATE_HEADER, userEstimate)
    If IsNumeric(userEstimate) Then
        Call TaskUtils.SetEbsEstimates(sheet, CDbl(userEstimate))
    Else
        'Clear the ebs list entries if invalid data was passed
        Dim br As Range
        Set br = TaskUtils.GetEbsListBodyRange(sheet)
        
        If Not br Is Nothing Then br.Delete
    End If
End Function



Function SetEbsEstimates(sheet As Worksheet, userEstimate As Double)
        'Populate the list of ebs estimates at the task sheet for the given user estimate.
        '
        'Input args:
        '  sheet: The task sheet one wants to add the data to
        '  userEstimate: The user estimate one wants to get propable finish times for by using the ebs algorithm
        
        'Clear list of ebs estimates first
        Dim br As Range
        Set br = TaskUtils.GetEbsListBodyRange(sheet)
        
        If Not br Is Nothing Then br.Delete
        
        'Read the name of the contributor to select the correct velocity pool later on
        Dim contributor As String
        contributor = Utils.GetSingleDataCellVal(sheet, Constants.CONTRIBUTOR_HEADER)
        
        If StrComp(contributor, "") = 0 Then Exit Function
        
        Dim contribPool() As Double
        contribPool = EbsUtils.GenVelocityPool(contributor, Constants.EBS_VELOCITY_POOL_SIZE)
        
        If Not Base.IsArrayAllocated(contribPool) Then Exit Function
        
        'Do ebs calculation for the given support points
        Dim monteCarloTimeEstimates() As Double
        monteCarloTimeEstimates = EbsUtils.GetMonteCarloTimeEstimates(userEstimate, contribPool, Constants.EBS_VELOCITY_PICKS)
        
        If Not Base.IsArrayAllocated(monteCarloTimeEstimates) Then Exit Function
                
        Dim supportPoints() As Double
        supportPoints = Utils.CopyVarArrToDoubleArr(Constants.EBS_SUPPORT_PROPABILITIES)
        
        Dim interpolatedEstimates() As Double
        interpolatedEstimates = EbsUtils.InterpolateEstimates(supportPoints, monteCarloTimeEstimates)
        
        'Now set the values to the sheet
        
        'Iterate over support points
        Dim rowIdx As Integer
        For rowIdx = 0 To UBound(supportPoints)
            Call TaskUtils.AddNewEbsEntry(sheet, CDbl(supportPoints(rowIdx)), interpolatedEstimates(rowIdx))
        Next rowIdx
End Function



Function GetEbsListBodyRange(sheet As Worksheet) As Range
    'Get the body range of the ebs list of a task sheet
    '
    'Input args:
    '   sheet:  The task sheet which
    '
    'Output args:
    '   GetEbsListBodyRange: The body range
    
    'Init output
    Set GetEbsListBodyRange = Nothing
    
    'Get the ebs data data table / list object
    Dim lo As ListObject
    Set lo = sheet.ListObjects(Constants.TASK_SHEET_EBS_LIST_IDX)
    
    If Not lo Is Nothing Then
        Set GetEbsListBodyRange = lo.DataBodyRange
    End If
End Function



Function GetTimeListBodyRange(sheet As Worksheet) As Range
    'Get the body range of the time tracking list of a task sheet
    '
    'Input args:
    '   sheet:  The task sheet which
    '
    'Output args:
    '   GetEbsListBodyRange: The body range
    
    'Init output
    Set GetTimeListBodyRange = Nothing
    
    'Get the ebs data data table / list object
    Dim lo As ListObject
    Set lo = sheet.ListObjects(Constants.TASK_SHEET_TIME_LIST_IDX)
    
    If Not lo.DataBodyRange Is Nothing Then
        Set GetTimeListBodyRange = lo.DataBodyRange
    End If
End Function



Function SetHash(sheet As Worksheet, hash As String)
    Call SetSingleDataCell(sheet, T_HASH_HEADER, hash)
End Function



Function SetKanbanList(sheet As Worksheet, list As String)
    Call Utils.SetSingleDataCell(sheet, KANBAN_LIST_HEADER, list)
End Function



Function SetContributor(sheet As Worksheet, contributor As String)
    Call Utils.SetSingleDataCell(sheet, CONTRIBUTOR_HEADER, contributor)
End Function



Function SetDueDate(sheet As Worksheet, datee As Date)
    Call Utils.SetSingleDataCell(sheet, Constants.DUE_DATE_HEADER, datee)
End Function



Function SetFinishedOnDate(sheet As Worksheet, datee As Date)
    Call Utils.SetSingleDataCell(sheet, Constants.TASK_FINISHED_ON_HEADER, datee)
End Function



Function SetComment(sheet As Worksheet, comment As String)
    Call Utils.SetSingleDataCell(sheet, COMMENT_HEADER, comment)
End Function



Function GetLastEndTimestamp(sheet As Worksheet) As Date
    'Returns the timestamp of the last row of tracked time entries
    '
    'Input args:
    '  sheet:  The task sheet to retrieve the val from
    
    'Init output
    GetLastEndTimestamp = CDate(0)
    
    Dim endStamps As Range
    Dim val As String
       
    Set endStamps = TaskUtils.GetTimeListColumn(sheet, Constants.END_TIME_HEADER, ceData)
    If Not endStamps Is Nothing Then
        'Read the last value from the list
        val = Utils.GetBottomRightCell(endStamps).Value
        If IsDate(val) Then
            GetLastEndTimestamp = CDate(Utils.GetBottomRightCell(endStamps).Value)
        End If
    End If
    
    'Debug info
    'Debug.Print endStamps.Address
End Function



Function GetVelocity(sheet As Worksheet) As Double
    'Wrapper to get the current veloctiy of the task (user estimated time over elapsed time)
    '2h user estimate over 1h elapsed time = 200% velocity
    'Set to -1 if invalid value was read
    '
    'Input args:
    '  sheet:  The task sheet to retrieve the val from
    
    Dim val As String
    val = GetSingleDataCellVal(sheet, Constants.VELOCITY_HEADER)
    If IsNumeric(val) Then
        GetVelocity = CDbl(val)
    Else
        GetVelocity = -1
    End If
End Function



Function GetEstimate(sheet As Worksheet) As Double
    'Wrapper to get the user estimate from the sheet (the time the user thought the task would take)
    '
    'Input args:
    '  sheet:  The task shete to retrieve the val from
    
    Dim val As String
    val = GetSingleDataCellVal(sheet, Constants.TASK_ESTIMATE_HEADER)
    If IsNumeric(val) Then
        GetEstimate = CDbl(val)
    Else
        GetEstimate = -1
    End If
End Function



Function GetHash(sheet As Worksheet) As String
    'Wrapper to get the hash from the sheet. "" if hash is invalid
    '
    'Input args:
    '  sheet:  The task shete to retrieve the val from
    
    Dim val As String
    val = GetSingleDataCellVal(sheet, Constants.T_HASH_HEADER)
    If SanityChecks.CheckHash(val) Then
        GetHash = val
    Else
        GetHash = ""
    End If
End Function



Function GetContributor(sheet As Worksheet) As String
    'Wrapper to get the contributor name from the sheet. "" if contributor name does not exist
    '
    'Input args:
    '  sheet:  The task shete to retrieve the val from
    '
    'Output args:
    '   GetContributor: The contributor name
    
    GetContributor = GetSingleDataCellVal(sheet, Constants.CONTRIBUTOR_HEADER)
End Function



Function GetTaskTotalTime(sheet As Worksheet) As Double
    'Wrapper to return the time that was spent on the task
    '
    'Input args:
    '  sheet: The task sheet
    'Output args:
    '  GetTaskTotalTime: The total time spent on the task (-1 if value could not be read)
    
    'Init output
    GetTaskTotalTime = -1
    
    'Check args
    If sheet Is Nothing Then Exit Function
    If Not SanityChecks.CheckHash(sheet.name) Then Exit Function
    
    Dim val As String
    val = Utils.GetSingleDataCellVal(sheet, Constants.TASK_TOTAL_TIME_HEADER)
    
    If IsNumeric(val) Then
        If CDbl(val) >= 0 Then
            'Valid value has been found. Return it
            GetTaskTotalTime = CDbl(val)
        End If
    End If
End Function



Function SetPlainDelta(clickedCell As Range)
    'Restore the 'standard' method to calculate the delta: Plain subtraction of two dates. Formula only works inside a list object
    '
    'Input args:
    '   clickedCell:    The cell a user clicked a hyperlink in. See TaskUtils.HandleFollowHyperlink
    
    Const PLAIN_DELTA_FORMULA As String = "=MAX(([@[" & Constants.END_TIME_HEADER & "]]-[@[" & Constants.START_TIME_HEADER & "]])*24,0)"
    Dim taskSheet As Worksheet
    Set taskSheet = clickedCell.parent
    
    'Get the cell to add the formula to
    Dim deltaTCell As Range
    Set deltaTCell = TaskUtils.IntersectTimeListColumn(taskSheet, Constants.TIME_DELTA_HEADER, clickedCell)
    
    If deltaTCell Is Nothing Then Exit Function
    
    deltaTCell.Formula = PLAIN_DELTA_FORMULA 'Be careful here to use english formula format (German ';' become ',') otherwise setting formula will fail
End Function



Function SetCalendarDelta(clickedCell As Range)
    'Function calculates the calendar appointment based time delta between two given dates of a time tracking row
    '
    'Input args:
    '   clickedCell:    The cell a user clicked a hyperlink in (see TaskUtils.HandleFollowHyperlink)
    
    Dim taskSheet As Worksheet
    Set taskSheet = clickedCell.parent
    Dim contributor As String
    
    contributor = TaskUtils.GetContributor(taskSheet)
    
    'if delta
    Dim startTimeCell As Range
    Dim endTimeCell As Range
    Dim deltaTCell As Range
    
    'Get cells of the time list to calculate delta from
    Set startTimeCell = TaskUtils.IntersectTimeListColumn(taskSheet, Constants.START_TIME_HEADER, clickedCell)
    Set endTimeCell = TaskUtils.IntersectTimeListColumn(taskSheet, Constants.END_TIME_HEADER, clickedCell)
    Set deltaTCell = TaskUtils.IntersectTimeListColumn(taskSheet, Constants.TIME_DELTA_HEADER, clickedCell)
    
    'Check cells and entries
    If startTimeCell Is Nothing Or endTimeCell Is Nothing Or deltaTCell Is Nothing Or StrComp(contributor, "") = 0 Then Exit Function
    If Not IsDate(startTimeCell.Value) Or Not IsDate(endTimeCell.Value) Then Exit Function
    
    Dim startTime As Date
    Dim endTime As Date
    
    startTime = CDate(startTimeCell.Value)
    endTime = CDate(endTimeCell.Value)
    
    'Read the outlook appointments
    Dim oItems As Outlook.Items
    Set oItems = CalendarUtils.GetCalItems(contributor, Constants.BUSY_AT_OPTIONAL_APPOINTMENTS)
    
    'Calculate delta time
    Dim deltaT As Double
    Dim appointment
    deltaT = CalendarUtils.MapDateToHours(contributor, oItems, endTime, startTime, _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOnset), _
        SettingUtils.GetContributorApptOnOffset(contributor, ceOffset))
        
    If deltaT <> -1 Then
        deltaTCell.Value = deltaT
    Else
        deltaTCell.Value = Constants.INVALID_ENTRY_PLACEHOLDER
    End If
End Function



Function SetClickActionHighlight(sheet As Worksheet, clickedCell As Range)
    'Highlight the clicked cell in a time list row of a task sheet to visualize the users click.
    'Clear all the accent colors from cells in the time list row which have clickable links. The clicked cell is highlighted afterwards
    '
    'Input args:
    '   sheet:  The task sheet
    '   clickedCell:    The cell the user clicked on
    
    Dim actionCells As Range
    'Get the two clickable cells
    
    'Add the first clickable cell of the row
    Set actionCells = TaskUtils.IntersectTimeListColumn(sheet, Constants.TASK_SHEET_ACTION_ONE_HEADER, clickedCell)
    
    'Add the second clickable cell of the row
    Set actionCells = Base.UnionN(actionCells, _
        TaskUtils.IntersectTimeListColumn(sheet, Constants.TASK_SHEET_ACTION_TWOO_HEADER, clickedCell))
    
    actionCells.Interior.color = xlNone
    
    'Highlight the clicked cell
    Dim accentColor As Long
    accentColor = SettingUtils.GetColors(ceAccentColor)
    
    clickedCell.Interior.color = accentColor
End Function



Function AddSpecialActionsToTimeListRow(sheet As Worksheet, rowIdentifier As Range)
    'Add extra actions to the task sheet. The actions can be catched by catching a hyperlink click. See 'TaskUtils.HandleFollowHyperlink'
    '
    'Input args:
    '   sheet:          The task sheet you want to add the action links to
    '   rowIdentifier:  Cell or row range that identifies the time list row you want to add the special functions for
    
    'Check args:
    If sheet Is Nothing Or rowIdentifier Is Nothing Then Exit Function
    
    Dim actionOneCell As Range
    Dim actionTwooCell As Range
    
    Set actionOneCell = Base.IntersectN(rowIdentifier.EntireRow, TaskUtils.GetTimeListColumn(sheet, Constants.TASK_SHEET_ACTION_ONE_HEADER, ceData))
    Set actionTwooCell = Base.IntersectN(rowIdentifier.EntireRow, TaskUtils.GetTimeListColumn(sheet, Constants.TASK_SHEET_ACTION_TWOO_HEADER, ceData))
    
    If Not actionOneCell Is Nothing And Not actionTwooCell Is Nothing Then
        'Add special action links to the list row to manage delta time calculation of task sheet row (by plain delta or with calendar data)
        'Add the function names
        actionOneCell.Value = Constants.TASK_SHEET_ACTION_ONE_NAME
        actionTwooCell.Value = Constants.TASK_SHEET_ACTION_TWOO_NAME
        
        'Add the hyperlinks which will be catched by an 'TaskUtils.HandleFollowHyperlink'
        Call Utils.AddSubtileHyperlink(actionOneCell, actionOneCell.Address)
        Call Utils.AddSubtileHyperlink(actionTwooCell, actionTwooCell.Address)
    End If
End Function
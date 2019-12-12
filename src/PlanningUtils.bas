Attribute VB_Name = "PlanningUtils"
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

Public Enum ShiftDirection
    ceShiftUp = 1
    ceShiftDown = -1
End Enum

Enum EstimateType
    ceTime = 0
    ceDate = 1
End Enum

Public Enum CopySource
    ceCopyFromPlanningList = 1
    ceCopyFromWorksheetData = -1
End Enum

Function GetPlanningSheet() As Worksheet
    'Return the main planning sheet of this project
    '
    'Output args:
    '  GetPlanningSheet:   A reference to the sheet with a fixed name
    
    Set GetPlanningSheet = ThisWorkbook.Worksheets(Constants.PLANNING_SHEET_NAME)
End Function



Function GetTaskList() As ListObject
    'Return the main list on the task sheet
    '
    'Output args:
    '  GetTaskList:    A reference to the list object with a fixed name
    Set GetTaskList = PlanningUtils.GetPlanningSheet.ListObjects(Constants.TASK_LIST_NAME)
End Function



Function GetTaskListColumn(colIdentifier As Variant, rowIdentifier As ListRowSelect) As Range
    'Wrapper to read column of a task list
    '
    'Input args:
    '  colIdentifier:  An identifier specifying the column one whishes to extract. Can be a cell inside the column or the header string
    '  rowIdentifier:  An identifier specifying whether to return the whole list, only the header or only the data range (without headers)
    '
    'Output args:
    '  GetTaskListColumn: Range of the selected cells of the column
    
    Set GetTaskListColumn = Utils.GetListColumn(GetPlanningSheet, Constants.TASK_LIST_NAME, colIdentifier, rowIdentifier)
End Function



Function AddNewTask(Optional ByRef newHash As String, Optional ByRef entryCell As Range)
    'Adds a new task and returns entry cell and generated hash value
    '
    'Input args:
    '  entryCell:  The cell a new task shall be added to. The cell will later contain the no. of the added task e.g. '#0001'
    '  newHash:    Hash value which will be set for the new task (if any)
    '
    'Output args:
    '  newHash:    If no hash value was given one can read the newly generated hash from here
    
    'Generate a new HASH if no hash was passed
    If Not SanityChecks.CheckHash(newHash) Then
        newHash = CreateHashString("t")
    End If
    
    'Add a new entry to task list'
    Dim entryAdded As Boolean
    entryAdded = False
    
    'Add the task line in planning sheet
    entryAdded = AddNewTaskLine(newHash, entryCell)
    
    If entryAdded Then
        'Create a new task sheet
        Call TaskUtils.GetTaskSheet(newHash)
        
        'Select the title of the created taskb
        Dim title As Range
        Set title = PlanningUtils.IntersectHashAndListColumn(newHash, Constants.TASK_NAME_HEADER)
        
        If Not title Is Nothing Then
            'Select title and manually invoke handling of selection changes (this is necessary if function is called with disabled events)
            title.Select
            Call Planning.HandleSelectionChanges(title)
        End If
    End If
End Function



Function AddNewTaskLine(hash As String, ByRef newEntryCell As Range) As Boolean
    'Adds a new task line in the planning sheet's main list
    '
    'Input args:
    '  hash:           Hash value which will be set for the new task (if any)
    '  newEntryCell:   The cell a new task shall be added to. The cell will later contain the no. of the added task e.g. '#0001'
    '
    'Output args:
    '  AddNewTaskLine: If everything went well result will be 'true'
    
    'Init output
    AddNewTaskLine = False
    
    Dim hashExists As Boolean
    hashExists = Not Utils.FindSheetCell(GetPlanningSheet, hash) Is Nothing
    
    'Check args
    If StrComp(hash, "") = 0 Or hashExists Then
        'Do not add a new entry if the hash already exists or if the hash is empty
        Exit Function
    End If
    
    Dim taskNameCell As Range
    Dim taskEstimateCell As Range
    Dim taskKanbanListCell As Range
    Dim taskContributorCell As Range
    Dim taskHashCell As Range
    Dim taskPrioCell As Range
    
    Dim newFormattedNumber As String
    Dim gotEntryData As Boolean
    
    'Get all the edited cells:
    'First get the new line the task will be added to
    gotEntryData = Utils.GetNewEntry(PlanningUtils.GetPlanningSheet, Constants.TASK_LIST_NAME, _
        newEntryCell, newFormattedNumber)
    Set taskNameCell = Utils.IntersectListColAndCells(GetPlanningSheet, Constants.TASK_LIST_NAME, TASK_NAME_HEADER, newEntryCell)
    Set taskEstimateCell = Utils.IntersectListColAndCells(GetPlanningSheet, Constants.TASK_LIST_NAME, TASK_ESTIMATE_HEADER, newEntryCell)
    Set taskKanbanListCell = Utils.IntersectListColAndCells(GetPlanningSheet, Constants.TASK_LIST_NAME, KANBAN_LIST_HEADER, newEntryCell)
    Set taskContributorCell = Utils.IntersectListColAndCells(GetPlanningSheet, Constants.TASK_LIST_NAME, CONTRIBUTOR_HEADER, newEntryCell)
    Set taskHashCell = Utils.IntersectListColAndCells(GetPlanningSheet, Constants.TASK_LIST_NAME, T_HASH_HEADER, newEntryCell)
    Set taskPrioCell = Utils.IntersectListColAndCells(GetPlanningSheet, Constants.TASK_LIST_NAME, Constants.TASK_PRIORITY_HEADER, newEntryCell)
    
    If newEntryCell Is Nothing Or _
        taskNameCell Is Nothing Or _
        taskEstimateCell Is Nothing Or _
        taskKanbanListCell Is Nothing Or _
        taskContributorCell Is Nothing Or _
        taskHashCell Is Nothing Or _
        taskPrioCell Is Nothing Then
        'Exit if a cell could not be found
        Exit Function
    End If
    
    'Add the generated hash to identify the task
    taskHashCell.Value = hash
    
    'Add number for new entry
    newEntryCell.Value = newFormattedNumber
    
    'Add placeholders for fields or initial values or const values
    taskNameCell.Value = Constants.TASK_NAME_INITIAL
    taskEstimateCell.Value = Constants.TASK_ESTIMATE_INITIAL
    taskKanbanListCell.Value = Constants.KANBAN_LIST_TODO
    taskContributorCell.Value = Constants.CONTRIBUTOR_INITIAL
    
    'Add a hyperlink to the sheet which will be generated in the next step
    Dim subAddress As String
    subAddress = Constants.TASK_SHEET_STD_LINK
    
    'Add a hyperlink with custom format
    Call Utils.AddSubtileHyperlink(taskNameCell, subAddress)
    
    'Do this as last step because priorities are resorting the table
    taskPrioCell.Value = Constants.TASK_PRIO_INITIAL
    
    'Call sorting prio column manually to since events should be disabled when calling 'AddNewTaskLine'
    Call PlanningUtils.OrganizePrioColumn
    
    'Search for the hash here as tasks can get reordered with priority sorting
    Set newEntryCell = Utils.FindSheetCell(GetPlanningSheet, hash)
    
    AddNewTaskLine = True
End Function



Function GetTaskHash(Optional ByRef rng As Range = Nothing) As String

    'Gets the task hash to a selection / cell range in a row. If a range arg is passed the passed range will be used to determine the
    'tasks hash. The rng arg is set by ref so that a current selection can also be retrieved if 'rng'is set to 'Nothing'when calling.
    'The selection is then checked against consisting of only one cell. If not the rng arg is set to 'Nothing'
    '
    'Input args:
    '  rng:            The cell you would like to get the task hash for
    '
    'Output args:
    '  GetTaskHash:    The hash string
    
    'Init the function
    GetTaskHash = ""
    
    If rng Is Nothing Then
        'Use selection if nothing was passed
        Set rng = Selection
    End If
    
    Dim activeSheetName As String
    Dim foundHashCell As Range
    Dim hashOfTask As Range
    
    'Check if we are in task sheet and reset range if condition does not apply
    If StrComp(rng.Parent.name, Constants.PLANNING_SHEET_NAME) <> 0 Then
        Debug.Print "Please give/select a range of the main sheet called '" & Constants.PLANNING_SHEET_NAME & "'."
        Set rng = Nothing
        Exit Function
    End If
    
    'Check if only one row is selected and reset range if condition does not apply
    If Not rng.Rows.Count = 1 Then
        Debug.Print "Please give/select a range with only one row."
        Set rng = Nothing
        Exit Function
    End If
    
    'Find the hash cell in task sheet - we need the column of this finding
    Set foundHashCell = Utils.IntersectListColAndCells(PlanningUtils.GetPlanningSheet(), Constants.TASK_LIST_NAME, Constants.T_HASH_HEADER, rng)
    If Not foundHashCell Is Nothing Then
        GetTaskHash = foundHashCell.Value
    Else
        GetTaskHash = ""
    End If
    
    'Print out debug info
    'Debug.Print "Hash of selected task is: " & hashOfTask
End Function



Function IsTaskTracking(entry As Variant) As Boolean
    'Finds out whether the task identified by a row cell is tracking the user's time
    '
    'Input args:
    '  entry:              The cell (Range) or hash (String) in a task's row for which one wants to know whether it is tracking
    '
    'Output args:
    '  IsSelectedTaskTracking: True when the indicator flag is set for the task
    
    'Init output
    IsTaskTracking = False

    'Get the currently selected valid task if any
    Dim taskHash As String
    
    Select Case TypeName(entry)
        Case "String"
            taskHash = entry
        Case "Range"
            Dim rng As Range
            Set rng = entry
            taskHash = PlanningUtils.GetTaskHash(rng)
    End Select
    
    If StrComp(taskHash, "") = 0 Then
        'No valid hash could be found
        IsTaskTracking = False
    Else
        'If the indicator sign is in the task row the task is tracking.
        Dim indicatorCell As Range
        Set indicatorCell = PlanningUtils.IntersectHashAndListColumn(taskHash, Constants.INDICATOR_HEADER)

        IsTaskTracking = StrComp(indicatorCell.Value, Constants.INDICATOR) = 0
    End If
End Function



Function EndAllTasks()
    'Ends the tracking of all tasks: Add finish timestamp, delete indicators
    
    Dim sheet As Worksheet

    'Set end times in all task sheets. Multiple started task should never exist. As precaution we always stop all tasks
    'in order to maintain a healthy document
    For Each sheet In Utils.GetAllTaskSheets()
        Call TaskUtils.SetEndTimeToSheetTracking(sheet)
    Next sheet

    'After setting end times to all task sheets we delete the '<current'tracker in the main sheet as well
    Dim indicatorCells As Range
    Set indicatorCells = Utils.FindSheetCell(PlanningUtils.GetPlanningSheet, Constants.INDICATOR)
    
    If Not indicatorCells Is Nothing Then
        indicatorCells.Value = ""
    End If
End Function



Function DeleteSelectedTask()
    'Function deletes the selected task: Its task row in planning sheet main list and its task sheet
    
    Dim hash As String
    Dim entryCell As Range
    Set entryCell = Nothing
    
    'We pass an empty entryCell here to get the selected cell.
    hash = PlanningUtils.GetTaskHash(entryCell)
    
    If Not SanityChecks.CheckHash(hash) Then
        Exit Function
    Else
        Call Utils.DeleteWorksheetSilently(TaskUtils.GetTaskSheet(hash))
        Call Utils.DeleteFilteredListObjectRow(entryCell)
    End If
End Function



Function StartSelectedTask()
    'Track the time for the selected task: Add a timestamp in task task sheet and set indicators in lists to see which task is active
    
    'End all task prior to start a new one
    Call EndAllTasks
    
    Dim hash As String
    Dim entryCell As Range
    Dim taskSheet As Worksheet
    Dim planningSheet As Worksheet
    
    Set entryCell = Nothing
    
    'We pass an empty entryCell here to retrieve the cell of the function.
    hash = PlanningUtils.GetTaskHash(entryCell)
    
    If Not SanityChecks.CheckHash(hash) Then
        Exit Function
    Else
        'Fetch the sheets
        Set taskSheet = TaskUtils.GetTaskSheet(hash)
        Set planningSheet = PlanningUtils.GetPlanningSheet
        
        'Add a new time entry for the selected task. Check if this was successful
        If TaskUtils.AddNewTrackingEntry(taskSheet) Then
        
            'Add a marker in the planning sheet
            Dim indicatorCell As Range
            Set indicatorCell = PlanningUtils.IntersectHashAndListColumn(hash, Constants.INDICATOR_HEADER)
            indicatorCell.Value = Constants.INDICATOR
            
            'If task is still marked as in backlog or done, mark it as in progress
            Dim kanbanListCell As Range
            Set kanbanListCell = PlanningUtils.IntersectHashAndListColumn(hash, Constants.KANBAN_LIST_HEADER)
            
            If StrComp(kanbanListCell.Value, Constants.KANBAN_LIST_TODO) = 0 Or _
                StrComp(kanbanListCell.Value, Constants.KANBAN_LIST_DONE) = 0 Then
                kanbanListCell.Value = Constants.KANBAN_LIST_IN_PROGRESS
                Call TaskUtils.SetKanbanList(taskSheet, Constants.KANBAN_LIST_IN_PROGRESS)
            End If
        End If
    End If
End Function



Function IntersectHashAndListColumn(hash As String, colIdentifier As Variant) As Range
    'Get a cell of a data column to a specific hash. Wrapper with hash cell search for list col intersection function
    '
    'Input args:
    '  hash:                       The hash of the task row one wants to intersect
    '  colIdentifier:              The column's identifier, can be a header string or a cell in the column one wants to intersect with
    '
    'Output args:
    '  IntersectHashAndListColumn: The intersecting cell
    
    'Init output
    Set IntersectHashAndListColumn = Nothing
    
    'Check args
    If Not SanityChecks.CheckHash(hash) Then Exit Function
    
    Dim hashCell As Range
    Set hashCell = PlanningUtils.GetTaskListColumn(Constants.T_HASH_HEADER, ceData)
    Set hashCell = Base.FindAll(hashCell, hash)
    Set IntersectHashAndListColumn = _
        Utils.IntersectListColAndCells(PlanningUtils.GetPlanningSheet, Constants.TASK_LIST_NAME, colIdentifier, hashCell)
End Function



Function ReinitPriorities()
    'This function takes all numeric entries in the priority column and applies 'new'priorities to them:
    'A priority list of [6, 4, 3, 2.3, 1] becomes [5, 4, 3, 2, 1]
    'The new list contains only adjacent whole numbers
    
    Dim priorities As Range
    Dim noPriorities As Range
    Dim hashes As Range
    Dim finishedTasks As Range: Set finishedTasks = PlanningUtils.GetFinishedTasks
    Dim unfinishedTasks As Range: Set unfinishedTasks = PlanningUtils.GetUnfinishedTasks
    
    If Not finishedTasks Is Nothing Then
        'Get task that do not have any priority - all finished tasks
        Set noPriorities = Base.IntersectN(finishedTasks.EntireRow, _
            PlanningUtils.GetTaskListColumn(Constants.TASK_PRIORITY_HEADER, ceData))
    End If
    
    If Not unfinishedTasks Is Nothing Then
        'Get taks that do have a priority - all unfinished tasks
        Set priorities = Base.IntersectN(unfinishedTasks.EntireRow, _
            PlanningUtils.GetTaskListColumn(Constants.TASK_PRIORITY_HEADER, ceData))
           
        'Intersect the unfinished task rows with the column that contains their hashes
        Set hashes = Base.IntersectN(unfinishedTasks.EntireRow, _
            PlanningUtils.GetTaskListColumn(Constants.T_HASH_HEADER, ceData))
    End If
    
    'Write N/A to all tasks with no priority
    If Not noPriorities Is Nothing Then noPriorities = Constants.N_A
    
    'Debug info
    'Debug.Print "Priorites address is: " + priorities.Address
    
    'Count cells with priorities
    
    If priorities Is Nothing Then Exit Function
    
    Dim prioCount As Long
    prioCount = priorities.Count
    
    If prioCount = 0 Then
        Exit Function
    End If
    
    'Use two arrays to store hashes of tasks and their priorities.
    'Sort them according to their priority and reinit the priorities with counting
    Dim cell As Range
    Dim prioIdx As Long
    Dim prioNumbers() As Double
    Dim hashStrings() As Variant
    
    ReDim prioNumbers(0 To prioCount - 1)
    ReDim hashStrings(0 To prioCount - 1)
    
    'Collect all array values
    prioIdx = 0
    For Each cell In priorities
        prioNumbers(prioIdx) = CDbl(cell.Value)
        prioIdx = prioIdx + 1
    Next cell
    
    prioIdx = 0
    For Each cell In hashes
        hashStrings(prioIdx) = cell.Value
        prioIdx = prioIdx + 1
    Next cell
    
    'Sort the priorities and the hashes
    Call QuickSort(prioNumbers, ceAscending, 0, UBound(prioNumbers), hashStrings)
    
    Dim newPrio As Long
    
    For newPrio = 0 To UBound(prioNumbers)
        'Reinit the prio
        Dim currentHash As String
        currentHash = hashStrings(newPrio)
    
        Dim prioCell As Range
        Set prioCell = IntersectHashAndListColumn(currentHash, Constants.TASK_PRIORITY_HEADER)
        If Not prioCell Is Nothing Then
            prioCell.Value = CDbl(newPrio + 1)
        End If
    Next newPrio
End Function



Function SetMultiCellHighlight(Optional sel As Range) As Range
    'Highlight multiple rows to show all entries with same column content
    '
    'Input args:
    '  sel:                    A passed selection one wants to set a highlight for. If on value is passed the current cursor selection will be evaluated
    '
    'Output args:
    '  SetMultiCellHighlight:  The range the highlight was set for (only for the multi selection enabled column, not the whole row range)
    
    'Fallback to cursor selection
    If IsMissing(sel) Then
        Set sel = Selection
    End If
    
    If sel.Count <> 1 Then
        Exit Function
    End If
    
    Dim val As String
    val = sel.Value
    
    Dim multiSel As Range
    Dim bdRange As Range
    Dim selRows As Range
    
    'Find all column cells containing the same value as the ref cell
    Set multiSel = PlanningUtils.GetTaskListColumn(sel, ceData)
    Set multiSel = Base.FindAll(multiSel, val)
    Set bdRange = PlanningUtils.GetTaskListBodyRange
    
    'Set highlight
    Dim accentColor As Long
    accentColor = SettingUtils.GetColors(ceAccentColor)
    
    If Not multiSel Is Nothing And Not bdRange Is Nothing Then
        If multiSel.Count > 1 Then
            'Apply accent color to all rows - only if more than one row is selected
            Set selRows = Base.IntersectN(bdRange, multiSel.EntireRow)
            selRows.Interior.color = accentColor
        End If
    End If
    
    'Return the selection of all cells containing the same value
    Set SetMultiCellHighlight = multiSel
    
    'Debug info
    'multiSel.Select
End Function



Function ResetHighlight()
    'Function resets the multi cell highlight. See also 'SetMultiCellHighlight'
    
    Dim lightColor As Long
    lightColor = SettingUtils.GetColors(ceLightColor)
    
    Dim dbc As Range
    Set dbc = PlanningUtils.GetTaskListBodyRange
    If Not dbc Is Nothing Then
        'Delete the accent color from the cells
        dbc.Interior.color = xlNone
    End If
End Function



Function GetAllListValidatedCols() As Range
    'Return cells of all list columns that use a validation with an automatically generated item list
    '(contributor column and all tag columns)
    '
    'Output args:
    '  GetAllListValidatedCols:    The range of the columns that are list validation enabled and automatically managed
    
    Dim allListValidatedCols As Range
    Dim listValidatedCol As Range
    
    'Collect all resetable cols according to defined tags of columns
    Dim tagHeaderCells As Range
    Set tagHeaderCells = PlanningUtils.GetTagHeaderCells
    
    If Not tagHeaderCells Is Nothing Then
        Dim tagCell As Range
        For Each tagCell In tagHeaderCells
            Set listValidatedCol = PlanningUtils.GetTaskListColumn(tagCell.Value, ceData)
            Set allListValidatedCols = Base.UnionN(allListValidatedCols, listValidatedCol)
        Next tagCell
    End If
    
    Set allListValidatedCols = Base.UnionN(allListValidatedCols, PlanningUtils.GetTaskListColumn(Constants.CONTRIBUTOR_HEADER, ceData))
    Set GetAllListValidatedCols = allListValidatedCols
End Function



Function CopyTask(newHash As String, cpSource As CopySource, Optional sourceListHash As String, Optional sourceSheet As Worksheet)
    'Copy a selected task with specific column values (its name, tags, priority, estimated time, comment, due date and contributor)
    '
    'Input args:
    '   cpSource:       Switch to select copy source from: Either planning sheet list or an (external) task sheet
    '   sourceListHash: The task of the hash that you want to copy. Mandatory if copying from planning list was selected
    '   sourceSheet:    The (external) sheet that holds data you want to copy. Mandatory if copying from source sheet was selected
    '   newHash:        The new hash that will be used for the copied task (given by value to trace the task)
    
    'Copy most of the fields of a given hash. Do not call this method with enabled events as it may have problematic consequences
        
    'Check args
    If Not SanityChecks.CheckHash(newHash) Or _
        (cpSource = CopySource.ceCopyFromPlanningList And Not SanityChecks.CheckHash(sourceListHash)) Or _
        (cpSource = CopySource.ceCopyFromWorksheetData And sourceSheet Is Nothing) Then
        Exit Function
    End If
    
    If Application.EnableEvents = True Then
        Debug.Print "Do not run this function with enabled events. Trouble ahead - exiting function."
        Exit Function
    End If
    
    Dim newEntry As Range
    
    'First add a new task and return the cell of the entry to have a reference for further copying
    Call PlanningUtils.AddNewTask(newHash, newEntry)
    'Debug.Print "New entry: " + newEntry.Address
    
    'Search for hash after adding the new task, because list items are reordered
    Dim existingEntry As Range
    Set existingEntry = Utils.FindSheetCell(GetPlanningSheet, sourceListHash)
    'Debug.Print "Existing entry: " + existingEntry.Address
    
    Dim copiedFields As Variant
    
    'Collect all the copied field headers
    copiedFields = Array(TASK_NAME_HEADER, TASK_PRIORITY_HEADER, TASK_ESTIMATE_HEADER, COMMENT_HEADER, DUE_DATE_HEADER, _
        CONTRIBUTOR_HEADER)
    
    'Add the tag headers as well (dynamically, because the user can use different tag column names)
    Dim header As Variant
    
    'Populate regex fields
    'Copy regex fields (tags)
    Dim tagHeaders As Range
    Set tagHeaders = PlanningUtils.GetTagHeaderCells
    
    If Not tagHeaders Is Nothing Then
        Dim tagHeader As Range
        For Each tagHeader In tagHeaders
            ReDim Preserve copiedFields(UBound(copiedFields) + 1)
            copiedFields(UBound(copiedFields)) = tagHeader
        Next tagHeader
    End If
    
    'Copy fields and handle changes if needed
    For Each header In copiedFields
        Dim unifiedHeader As String
        unifiedHeader = Planning.UnifyTagName(CStr(header))
        
        Dim existingVal As Variant: existingVal = ""
        
        Select Case cpSource
            Case CopySource.ceCopyFromPlanningList
                existingVal = PlanningUtils.IntersectHashAndListColumn(sourceListHash, CStr(header)).Value
                
            Case CopySource.ceCopyFromWorksheetData
                Select Case unifiedHeader
                    Case TAG_REGEX
                        Dim readSerTagHeaders As String
                        Dim readSerTagValues As String
                        readSerTagHeaders = Utils.GetSingleDataCellVal(sourceSheet, Constants.SERIALIZED_TAGS_HEADERS_HEADER)
                        readSerTagValues = Utils.GetSingleDataCellVal(sourceSheet, Constants.SERIALIZED_TAGS_VALUES_HEADER)
                        
                        Dim readTagHeaders() As String
                        Dim readTagValues() As String
                        
                        readTagHeaders = Utils.CopyVarArrToStringArr(Utils.DeserializeArray(readSerTagHeaders))
                        readTagValues = Utils.CopyVarArrToStringArr(Utils.DeserializeArray(readSerTagValues))
                        
                        If Base.IsArrayAllocated(readTagHeaders) And Base.IsArrayAllocated(readTagValues) Then
                            Dim headIdx As Integer
                            For headIdx = 0 To UBound(readTagHeaders)
                                If StrComp(readTagHeaders(headIdx), header) = 0 Then
                                    existingVal = readTagValues(headIdx)
                                    headIdx = UBound(readTagHeaders)
                                End If
                            Next headIdx
                        End If
                        
                    Case Constants.TASK_PRIORITY_HEADER
                        'Do nothing here when copying based on worksheet data
                        existingVal = Constants.TASK_PRIO_INITIAL
                    
                    Case Else
                        existingVal = Utils.GetSingleDataCellVal(sourceSheet, CStr(header))
                End Select
        End Select
        
        Dim cell As Range
        Set cell = PlanningUtils.IntersectHashAndListColumn(newHash, CStr(header))
        
        'Handle cell value changes here
        Select Case unifiedHeader
            Case TASK_NAME_HEADER
                'Copy name and handle name change to copy the name to the task sheet
                existingVal = existingVal + " (copy/import)"
                cell.Value = existingVal
                Call Planning.HandleChanges(cell)
                
            Case TASK_PRIORITY_HEADER, TASK_ESTIMATE_HEADER, COMMENT_HEADER, DUE_DATE_HEADER
                'Copy values and handle change to copy values to task sheet
                cell.Value = existingVal
                Call Planning.HandleChanges(cell)
                
            Case Constants.TAG_REGEX
                'Copy the tag
                cell.Value = existingVal
                Call Planning.ManageTagChange(cell, False)
                
            Case Constants.CONTRIBUTOR_HEADER
                'Handle contributor change without cell validation update. This causes cell validation to fail
                cell.Value = existingVal
                Call Planning.ManageContributorChange(cell, False)
        End Select
    Next header
    
    'At the end sort the tasks with their priorities
    PlanningUtils.OrganizePrioColumn
End Function



Function ShiftPrio(dir As ShiftDirection, Optional rng As Range)
    'Change a tasks priority combined with sorting the priorities to shift tasks up and down
    '
    'Input args:
    '  dir: Direction to which to shift the task (up or down)
    '  rng: A passed range identifiying the shifted task. If nothing is passed use the current selection
    
    If IsMissing(rng) Or rng Is Nothing Then
        Set rng = Selection
    End If
    
    Dim startHash As String
    Dim startColHeader As String
    startHash = PlanningUtils.GetTaskHash(Selection)
    startColHeader = Utils.GetListColumnHeader(Selection)
    
    'Check if only one col is selected
    If Not rng.Columns.Count = 1 Then
        Debug.Print "Please give/select a range with only one col."
        rng = Nothing
        Exit Function
    End If
    
    'Get the cell containing the priority of the task
    Set rng = Utils.IntersectListColAndCells(GetPlanningSheet(), Constants.TASK_LIST_NAME, Constants.TASK_PRIORITY_HEADER, rng)
    
    If rng Is Nothing Then
        Exit Function
    End If
    
    Dim area As Range
    Dim cellAreaCount As Long
    Dim idx As Long
    Dim cell As Range
    Dim newVal As Double
        
    'Now depending on in which direction the task will be shifted do the following:
    '(1) Get all the selection areas (ranges that stick together)
    '(2) Shift every area individually:
    '(3) Get the top or bottom neighbour of the area if you are shifting downwards or the top neighbour if you are
    '   Shifting upwards. Then raise the prio value of the area to be slightly lower (shifting downwards) or higher (shifting upwards) as the
    '   neighbour priority
    '(4) After changing the priorities reinit them and resort them by calling the 'PlanningUtils.OrganizePrioColumn'function
    
    Select Case dir
        Case ShiftDirection.ceShiftUp
            For Each area In rng.areas
                Dim topNeighbour As Range
                Set topNeighbour = Utils.GetTopNeighbour(area)
                              
                If IsNumeric(topNeighbour.Value) Then
                    cellAreaCount = area.cells.Count
                
                    For idx = 1 To cellAreaCount
                        'Calculate a value slightly higher than the top neighbour priority and maintain the original order of the tasks
                        newVal = topNeighbour.Value + (cellAreaCount - idx + 1) * 0.001
                        area.cells(idx) = newVal
                    Next idx
                End If
            Next area
        
        Case ShiftDirection.ceShiftDown
            For Each area In rng.areas
                Dim bottomNeighbour As Range
                Set bottomNeighbour = Utils.GetBottomNeighbour(area)
                              
                If IsNumeric(bottomNeighbour.Value) Then
                    cellAreaCount = area.cells.Count
                
                    For idx = 1 To cellAreaCount
                        'Calculate a value slightly lower than the top neighbour priority and maintain the original order of the tasks
                        newVal = bottomNeighbour.Value - (idx + 1) * 0.001
                        area.cells(idx) = newVal
                    Next idx
                End If
            Next area
    End Select
    
    'If you break here you can see the changed values in the spreadsheet
    PlanningUtils.OrganizePrioColumn
    
    'Now select the hash where you started with
    Dim newSel As Range
    Set newSel = PlanningUtils.IntersectHashAndListColumn(startHash, startColHeader)
    
    If Not newSel Is Nothing Then
        newSel.Select
        Call Planning.HandleSelectionChanges(newSel)
    End If
    
End Function



Function GatherTasks(Optional rng As Range)
    'Gather tasks in priority queue
    '
    'Input args:
    '  rng: A passed range identifiying the tasks that shall be gathered. If nothing is passed use the current selection
    
    If IsMissing(rng) Or rng Is Nothing Then
        Set rng = Selection
    End If
    
    If rng.Count <= 1 Then
        'It is not useful to gather one single tasks around itself
        Exit Function
    End If
    
    Dim startHash As String
    Dim startColHeader As String
    startHash = PlanningUtils.GetTaskHash(Selection)
    startColHeader = Utils.GetListColumnHeader(Selection)
    
    'Check if only one col is selected
    If Not rng.Columns.Count = 1 Then
        Debug.Print "Please give/select a range with only one col."
        rng = Nothing
        Exit Function
    End If
    
    'Get the cell containing the priority of the task
    Set rng = Utils.IntersectListColAndCells(GetPlanningSheet(), Constants.TASK_LIST_NAME, Constants.TASK_PRIORITY_HEADER, rng)
    
    If rng Is Nothing Then
        Exit Function
    End If
    
    Dim cell As Range
    Dim idx As Long
    Dim newVal As Double
        
    'Then raise the prio value of the area to be slightly lower (shifting downwards) or higher (shifting upwards) as the
    '   neighbour priority
    'After changing the priorities reinit them and resort them by calling the 'PlanningUtils.OrganizePrioColumn'function
    Dim gatherVal As Double
    
    Dim gatherCell As Range
    Set gatherCell = PlanningUtils.IntersectHashAndListColumn(startHash, Constants.TASK_PRIORITY_HEADER)
    
    gatherVal = gatherCell.Value
    
    idx = 0
    For Each cell In rng
        'Calculate a value slightly lower than the priority the tasks shall be gathered around. Maintain the original order of the tasks
        newVal = gatherVal - idx * 0.001
        cell.Value = newVal
        idx = idx + 1
    Next cell
        
    'If you break here you can see the changed values in the spreadsheet
    PlanningUtils.OrganizePrioColumn
    
    'Now select the hash where you started with
    Dim newSel As Range
    Set newSel = PlanningUtils.IntersectHashAndListColumn(startHash, startColHeader)
    
    If Not newSel Is Nothing Then
        newSel.Select
        Call Planning.HandleSelectionChanges(newSel)
    End If
    
End Function



Function OrganizePrioColumn()
    'This function organizes the priorities in the priority column so that tasks keep their original priority order but with
    'a cleaner 'look'. First all priorities will be replaced with descending, whole numbers. Then the column is resorted descending
    'High priority values get to the top and low value to the bottom
    
    'First step: Clean values
    Call PlanningUtils.ReinitPriorities
    
    'Second step: Resort
    Dim lo As ListObject
    
    Set lo = PlanningUtils.GetTaskList
    Dim prioCol As Range
    Set prioCol = PlanningUtils.GetTaskListColumn(Constants.TASK_PRIORITY_HEADER, ceAll)
    
    Dim sf As SortFields
    Set sf = lo.Sort.SortFields
    sf.Clear
    sf.Add Key:=prioCol, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    With lo.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Function



Function GetCumulativeMode(ebsColCell As Range) As CumulativeMode
    'Function is used to read the setting value of an ebs column in task sheet.
    'Mode can either be cumulative or single value. This specifies if the times and dates displayed are showing the time order (cumulative times)
    'or the times for a single task to finish.
    '
    'Input args:
    '  ebsColCell: Any cell of the ebs column one whishes read the cumulative mode for
    '
    'Output args:
    '  GetCumulativeMode:
    
    'Init output
    GetCumulativeMode = CumulativeMode.ceInvalid
    
    'Check args
    If ebsColCell Is Nothing Then Exit Function
    
    Dim headerCell As Range
    Dim cumulativeSetCell As Range
    
    Set headerCell = PlanningUtils.GetTaskListColumn(ebsColCell, ceHeader)
    Set cumulativeSetCell = Utils.GetTopNeighbour(headerCell)
    
    Select Case cumulativeSetCell.Value
        Case "single"
            GetCumulativeMode = CumulativeMode.ceSingle
        Case "cumulative"
            GetCumulativeMode = CumulativeMode.ceCumulative
        End Select
End Function



Function CollectEbsColData(headerCell As Range)
    'The function reads data from ebs sheet and copies it to the planning sheet.
    'Settings in the column header (propability of estimate and time or date mode) are taken into account
    
    Dim header As String
    header = headerCell.Value
    
    Dim regex As New RegExp
    regex.Global = True
    
    regex.Pattern = Constants.EBS_COLUMN_REGEX
    
    If regex.test(header) Then
        'If the header text of the cell is an ebs col header: read the col propability of the estimates to be displayed
        Dim matches As MatchCollection
        Set matches = regex.Execute(header)
        
        Dim percentage As Double
        If matches.Count = 1 Then
            If matches.item(0).SubMatches.Count > 0 Then
                percentage = CInt(matches.item(0).SubMatches.item(0)) / 100
                'Debug info
                'Debug.Print "Ebs col percentage is: " & Format(percentage, "0#%")
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
          
        'Get all the finished and unfinished tasks and init them with 'N/A'in the column
        Dim finishedTasks As Range
        Dim unfinishedTasks As Range
                
        Set finishedTasks = finishedTasks
        Set finishedTasks = Utils.IntersectListColAndCells(PlanningUtils.GetPlanningSheet, Constants.TASK_LIST_NAME, header, finishedTasks)
        
        Set unfinishedTasks = PlanningUtils.GetUnfinishedTasks
        Set unfinishedTasks = Utils.IntersectListColAndCells(PlanningUtils.GetPlanningSheet, Constants.TASK_LIST_NAME, header, unfinishedTasks)
        
        If Not finishedTasks Is Nothing Then
            finishedTasks = Constants.N_A
        End If
        
        If unfinishedTasks Is Nothing Then Exit Function
        unfinishedTasks = Constants.N_A
        
        'Find out whether the column shall display cumulative data or not.
        Dim cumuMode As CumulativeMode
        cumuMode = PlanningUtils.GetCumulativeMode(headerCell)
        'Debug info
        'Debug.Print "Cumulative mode is " & cumuMode
        
        'Find out whether time (h) or date shall be displayed
        Dim estType As EstimateType
        
        If InStr(1, header, "time") Then
            estType = EstimateType.ceTime
        ElseIf InStr(1, header, "date") Then
            estType = EstimateType.ceDate
        End If
        
        Dim estHeader As String
        
        'Decide which header to use to get data from the ebs sheet
        Select Case estType
            Case EstimateType.ceTime
                estHeader = Constants.EBS_RUNDATA_INTERPOLATED_TIME_HEADER
            Case EstimateType.ceDate
                estHeader = Constants.EBS_RUNDATA_INTERPOLATED_DATES_HEADER
        End Select
        
        Dim hash As String
        Dim contributor As String
                
        Dim cell As Range
        Dim contribSheet As Worksheet
        Dim tHashCol As Range
        Dim ebsRunDataHashCell As Range
        Dim estimateCell As Range
        Dim supportPointsCell As Range
        Dim ebsHashRow As Range
        Dim estVal As String
        Dim supportPoints() As Double
        Dim interpolatedVal() As Double
        Dim estArray() As Double
        
        Select Case cumuMode
            Case CumulativeMode.ceCumulative
                
                For Each cell In unfinishedTasks
                    'Cycle through all unfinished tasks and get cumulative estimates for them. The estimates are only valid in the
                    'queued order of the last ebs run. If the user reordered the tasks after the current ebs run one has to rerun this method
                    'to get the current estimate values. The method is therefore triggered after each ebs run.
                    hash = PlanningUtils.GetTaskHash(cell)
                    
                    If SanityChecks.CheckHash(hash) Then
                        'Read the data for the current contributor: First get ebs sheet of contributor
                        contributor = PlanningUtils.IntersectHashAndListColumn(hash, Constants.CONTRIBUTOR_HEADER)
                        Set contribSheet = EbsUtils.GetEbsSheet(contributor)
                        If contribSheet Is Nothing Then GoTo r2yNextCell
                        
                        'Get all the hashes
                        Set tHashCol = EbsUtils.GetRunDataListColumn(contribSheet, Constants.T_HASH_HEADER, ceData)
                        If tHashCol Is Nothing Then GoTo r2yNextCell
                        
                        'Get the rundata entry in ebs sheet with the current hash
                        Set ebsRunDataHashCell = Base.FindAll(tHashCol, hash)
                        If ebsRunDataHashCell Is Nothing Then GoTo r2yNextCell
                        
                        'Intersect to get the cell with the stored estimate array in it.
                        Set estimateCell = EbsUtils.IntersectRunDataColumn(contribSheet, estHeader, ebsRunDataHashCell)
                        If estimateCell Is Nothing Then GoTo r2yNextCell
                        
                        'Intersect to get the support points cell as well
                        Set supportPointsCell = EbsUtils.IntersectRunDataColumn(contribSheet, Constants.EBS_SUPPORT_POINT_HEADER, ebsRunDataHashCell)
                        If supportPointsCell Is Nothing Then GoTo r2yNextCell
                        
                        'Deserialize support points and estimates, interpolated for the whished value
                        supportPoints = Utils.CopyVarArrToDoubleArr(Utils.DeserializeArray(supportPointsCell.Value))
                        If Not Base.IsArrayAllocated(supportPoints) Then GoTo r2yNextCell
                        
                        Select Case estType
                            Case EstimateType.ceTime
                                estArray = Utils.CopyVarArrToDoubleArr(Utils.DeserializeArray(estimateCell.Value))
                            Case EstimateType.ceDate
                                'Cascade array casting as date strings cannot be converted to double directly
                                estArray = Utils.CopyVarArrToDoubleArr(Utils.CopyVarArrToDateArr(Utils.DeserializeArray(estimateCell.Value)))
                        End Select
                        
                        interpolatedVal = Utils.InterpolateArray(supportPoints, estArray, percentage)
                        
                        Select Case estType
                            Case EstimateType.ceTime
                                'Use this number format to crop the decimal places
                                cell.NumberFormat = "0.00"
                                'Only one value was interpolated and returned in the array.
                                cell.Value = interpolatedVal(0)
                            Case EstimateType.ceDate
                                'General format will display date correctly, reset the format. Otherwise date will be displayed in numbers (650000,3)
                                cell.NumberFormat = "General"
                                cell.Value = CDate(interpolatedVal(0))
                        End Select
                    End If
r2yNextCell:
                Next cell

            Case CumulativeMode.ceSingle
                'For non-cumulative values the task sheets data will be read. The data is generated when the user enters an estimate in the
                'planning sheet.
                'Cumulative data instead (see above) will be read from the ebs sheet of the contributor. So the data sources are very different -
                'With non-cumulative (single) data one can compare the user estimate, the real used time to finish the task and the ebs estimate.
                'This will make evaluation of the ebs algorithm possible.
                
                'Get the data from all! tasks
                
                Dim cycleRange As Range
                Set cycleRange = Base.UnionN(finishedTasks, unfinishedTasks)
                                    
                If estType = EstimateType.ceDate Then
                    'For non-cumulative mode only time values are allowed. Calculating date values would be possible but is not useful as tasks can only
                    'finished one after another
                    Exit Function
                End If
                
                For Each cell In cycleRange
                    hash = PlanningUtils.GetTaskHash(cell)
                    
                    'Read the task sheet
                    Dim taskSheet As Worksheet
                    Set taskSheet = TaskUtils.GetTaskSheet(hash)
                    
                    'Actually multiple cells are read here and converted to an array afterwards
                    Set supportPointsCell = TaskUtils.GetEbsListColumn(taskSheet, Constants.SINGLE_SUPPORT_POINT_HEADER, ceData)
                    Set estimateCell = TaskUtils.GetEbsListColumn(taskSheet, Constants.EBS_SELF_TIME_HEADER, ceData)
                    
                    If supportPointsCell Is Nothing Or estimateCell Is Nothing Then
                        GoTo g5hNextCell
                    End If
                    
                    supportPoints = Utils.CopyVarArrToDoubleArr(supportPointsCell.Value)
                    'Time est array (h) is converted here
                    estArray = Utils.CopyVarArrToDoubleArr(estimateCell.Value)
                    
                    If Not Base.IsArrayAllocated(supportPoints) Or Not Base.IsArrayAllocated(estArray) Then
                        GoTo g5hNextCell
                    End If
                    
                    'Now interpolate
                    interpolatedVal = Utils.InterpolateArray(supportPoints, estArray, percentage)
                    
                    cell.NumberFormat = "0.00"
                    cell.Value = interpolatedVal
g5hNextCell:
                Next cell
                
            Case CumulativeMode.ceInvalid
                Exit Function
        End Select
    End If
End Function



Function GetUnfinishedTasks() As Range
    'Return all the unfinished task cells (column of kanban list header)
    'Unfinished tasks are the 'negative'of finished tasks. Finished tasks can be determined with 'Done'label, unfinished tasks can have
    'various different labels
    
    'Init output
    Set GetUnfinishedTasks = Nothing
    
    Dim kanbanData As Range
    Dim finishedTasks As Range
    Dim unfinishedTasks As Range
    
    Set kanbanData = PlanningUtils.GetTaskListColumn(Constants.KANBAN_LIST_HEADER, ceData)
    
    If kanbanData Is Nothing Then Exit Function
    Set finishedTasks = Base.FindAll(kanbanData, Constants.KANBAN_LIST_DONE)
    
    'Remove finishedTasks from all tasks (if any)
    Set GetUnfinishedTasks = Base.Difference(kanbanData, finishedTasks)
End Function



Function GetFinishedTasks() As Range
    'Use the kanban list tag 'Done'to find all finished tasks
    
    'Init output
    Set GetFinishedTasks = Nothing
    
    Dim kanbanData As Range
    Dim finishedTasks As Range
    
    Set kanbanData = PlanningUtils.GetTaskListColumn(Constants.KANBAN_LIST_HEADER, ceData)
    If kanbanData Is Nothing Then Exit Function
    Set GetFinishedTasks = Base.FindAll(kanbanData, Constants.KANBAN_LIST_DONE)
    'Debug info
    'Debug.Print "Kanban data: " & kanbanData.Address
End Function



Function CollectTotalTimesSpent()
    'For every task on planning sheet collect the total time the user has spent on it
    Dim timeCells As Range
    Set timeCells = PlanningUtils.GetTaskListColumn(Constants.TOTAL_TIME_HEADER, ceData)
    
    If timeCells Is Nothing Then Exit Function
    
    Dim cell As Range
    Dim hash As String
    Dim sheet As Worksheet
    Dim totalTime
    
    For Each cell In timeCells
        'Cycle through all the cells the time will be stored in later on. Read the total time from the task sheet
        hash = PlanningUtils.GetTaskHash(cell)
        
        If SanityChecks.CheckHash(hash) Then
            Set sheet = TaskUtils.GetTaskSheet(hash)
            totalTime = TaskUtils.GetTaskTotalTime(sheet)
            cell.Value = totalTime
        End If
    Next cell
End Function



Function UpdateAllEbsCols()
    'Find all ebs cols on planning sheet and call the update function for them
    
    Dim colHeaderCells As Range
    Set colHeaderCells = Utils.FindSheetCell(PlanningUtils.GetPlanningSheet, Constants.EBS_COLUMN_REGEX, ceRegex)
    Dim headerCell As Range
    
    For Each headerCell In colHeaderCells
        Call PlanningUtils.CollectEbsColData(headerCell)
    Next headerCell
End Function



Function GetTaskListBodyRange() As Range
    'Get the body range of the task list. Can be used for intersection.
    
    'Init output
    Set GetTaskListBodyRange = Nothing
    
    Dim lo As ListObject
    Set lo = PlanningUtils.GetPlanningSheet.ListObjects(Constants.TASK_LIST_NAME)
    If Not lo Is Nothing Then
        Dim dbc As Range
        Set dbc = lo.DataBodyRange
        Set GetTaskListBodyRange = dbc
    End If
End Function



Function FollowTaskSheetLink(Target As hyperlink)
    'Follows the link to a task sheet unhiding it prior to selecting it
    
    If StrComp(Target.subAddress, Constants.TASK_SHEET_STD_LINK) = 0 Then
        Dim taskSheet As Worksheet
        Set taskSheet = TaskUtils.GetTaskSheet(PlanningUtils.GetTaskHash(Target.Parent))
        
        If Not taskSheet Is Nothing Then
            taskSheet.Visible = XlSheetVisibility.xlSheetVisible
        End If
        
        'After sheet is visible select sheet
        taskSheet.Select
    End If
End Function



Function GetTagHeaderCells() As Range
    Set GetTagHeaderCells = Utils.FindSheetCell(GetPlanningSheet, Constants.TAG_REGEX, ComparisonTypes.ceRegex)
End Function



Function GetSerializedTags(hash As String, Optional ByRef serializedTagHeaders As String) As String
    'Check args
    If StrComp(hash, "") = 0 Then Exit Function
    
    'Init output
    GetSerializedTags = ""
    
    Dim headers As Range
    Set headers = PlanningUtils.GetTagHeaderCells
    
    If Not headers Is Nothing Then
        Dim coV As New Collection
        Dim coH As New Collection
        Dim tagHeader As Range
        Dim tag As Range
        
        For Each tagHeader In headers
            Set tag = PlanningUtils.IntersectHashAndListColumn(hash, tagHeader)
            
            If Not tag Is Nothing Then
                If StrComp(tag, "") <> 0 Then
                    Call coV.Add(tag.Value)
                    Call coH.Add(tagHeader.Value)
                End If
            End If
        Next tagHeader
        
        If coV.Count > 0 Then
            'Serialize tags
            GetSerializedTags = Utils.SerializeArray(Base.CollectionToArray(coV))
            serializedTagHeaders = Utils.SerializeArray(Base.CollectionToArray(coH))
        End If
    End If
End Function

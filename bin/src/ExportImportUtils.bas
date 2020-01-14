Attribute VB_Name = "ExportImportUtils"
'  This macro collection lets you organize your tasks and schedules
'  for you with the evidence based design (EBS) approach by Joel Spolsky.
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

Const EXCHANGE_WORKBOOK_PREFIX As String = "EbsExportImport_"

Option Explicit



Function ExportVisibleTasks()
    'This function exports all sheets of tasks that are visible in the planning sheet list.
    'Tasks are exported to a separate workbook only contaning the sheets of tasks
    Dim hashRange As Range
    
    Set hashRange = PlanningUtils.GetTaskListColumn(Constants.T_HASH_HEADER, ceData)
    Dim cll As Range
    Dim visibleHashes As Range
    
    'Detect all visible tasks
    For Each cll In hashRange
        If cll.Height <> 0 Then
            'If the cell height is not 0 the cell is visible
            Set visibleHashes = Base.UnionN(visibleHashes, cll)
        End If
    Next cll
        
    'Debug info
    'Debug.Print "Visible hashes: " & visibleHashes.Count & " out of " & hashRange.Count
    
    If Not visibleHashes Is Nothing Then
        'In case there are visible hashes store them to a new workbook.
        Dim exportWb As Workbook
        Set exportWb = Excel.Workbooks.Add
        
        Dim firstWorksheet As Worksheet
        Set firstWorksheet = exportWb.Worksheets(1)
        
        Dim sheet As Worksheet
                    
        For Each cll In visibleHashes
            Set sheet = TaskUtils.GetTaskSheet(cll.Value)
            
            Dim copiedSheet As Worksheet
            If Not sheet Is Nothing Then
                Call sheet.Copy(after:=firstWorksheet)
            End If
        Next cll
        
        'Make sure the sheets are visible
        For Each sheet In exportWb.Worksheets
            sheet.Visible = xlSheetVisible
        Next sheet
        
        'The first sheet of the workbook is empty - delete it to only have task sheets inside the workbook
        Call Utils.DeleteWorksheetSilently(firstWorksheet)
    End If
End Function



Function ImportTasks(sourceWb As Workbook)
    'This function imports all task sheet of a workbook
    '
    'Input args
    '   sourceWb:   The workbook from which tasks are imported. The function reads all task sheets of the given workbook (identified by hash)
    'Check args
    
    If sourceWb Is Nothing Then Exit Function
    
    Dim sheet As Worksheet
    
    For Each sheet In sourceWb.Worksheets
        Dim cpHash As String
        cpHash = sheet.name
        If SanityChecks.CheckHash(cpHash) Then
            'If the sheet is a task sheet then copy the task from sheet data
            Dim newHash As String
            newHash = CreateHashString("t")
            
            Call ExportImportUtils.ImportSingleTask(newHash:=newHash, sourceSheet:=sheet)
        End If
    Next sheet
End Function



Function ImportSingleTask(newHash As String, sourceSheet As Worksheet)
    'Import a task from a task sheet. This function is very similar to the 'PlanningUtils.CopyTask' function but uses sheet data instead of
    'planning list data to 'duplicate' a task
    '
    'Input args:
    '   sourceSheet:    The (external) sheet that holds data you want to copy
    '   newHash:        The new hash that will be used for the copied task (given by value to trace the task)
    
    'Do not call this method with enabled events as it may have problematic consequences
        
    'Check args
    If Not SanityChecks.CheckHash(newHash) Or sourceSheet Is Nothing Then Exit Function
    
    If Application.EnableEvents = True Then
        Debug.Print "Do not run this function with enabled events. Trouble ahead - exiting function."
        Exit Function
    End If
    
    'First add a new task and return the cell of the entry to have a reference for further copying
    Call PlanningUtils.AddNewTask(newHash)
    
    Dim newTaskSheet As Worksheet
    Set newTaskSheet = TaskUtils.GetTaskSheet(newHash)
    
    Dim sourceTimeLo As Range
    Dim destTimeLo As Range
    
    Set sourceTimeLo = sourceSheet.ListObjects(Constants.TASK_SHEET_TIME_LIST_IDX).Range
    Set destTimeLo = newTaskSheet.ListObjects(Constants.TASK_SHEET_TIME_LIST_IDX).Range
        
    Dim sourceEbsLo As Range
    Dim destEbsLo As Range
    
    Set sourceEbsLo = sourceSheet.ListObjects(Constants.TASK_SHEET_EBS_LIST_IDX).Range
    Set destEbsLo = newTaskSheet.ListObjects(Constants.TASK_SHEET_EBS_LIST_IDX).Range
    
    Call sourceTimeLo.Copy
    Call destTimeLo.PasteSpecial(xlPasteAll)
    
    Call sourceEbsLo.Copy
    Call destEbsLo.PasteSpecial(xlPasteAll)
    
    Dim copiedFields As Variant
    
    'Collect all the copied field headers
    copiedFields = Array(TASK_NAME_HEADER, TASK_PRIORITY_HEADER, TASK_ESTIMATE_HEADER, KANBAN_LIST_HEADER, COMMENT_HEADER, DUE_DATE_HEADER, _
        CONTRIBUTOR_HEADER, TASK_FINISHED_ON_HEADER)
    
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
        
        Select Case unifiedHeader
            Case TAG_REGEX
                'Tags are stored in a serialized string inside the sheet. Deserialize the values and get the tag corresponding to
                'the current tag header
                        
                Dim readSerTagHeaders As String
                Dim readSerTagValues As String
                readSerTagHeaders = Utils.GetSingleDataCellVal(sourceSheet, Constants.SERIALIZED_TAGS_HEADERS_HEADER)
                readSerTagValues = Utils.GetSingleDataCellVal(sourceSheet, Constants.SERIALIZED_TAGS_VALUES_HEADER)
                        
                Dim readTagHeaders() As String
                Dim readTagValues() As String
                        
                readTagHeaders = Utils.CopyVarArrToStringArr(Utils.DeserializeArray(readSerTagHeaders))
                readTagValues = Utils.CopyVarArrToStringArr(Utils.DeserializeArray(readSerTagValues))
                        
                If Base.IsArrayAllocated(readTagHeaders) And Base.IsArrayAllocated(readTagValues) Then
                    'Values were deserialized successfully
                    Dim headIdx As Integer
                    For headIdx = 0 To UBound(readTagHeaders)
                        If StrComp(readTagHeaders(headIdx), header) = 0 Then
                            'Tag for current header was found.
                            existingVal = readTagValues(headIdx)
                                    
                            'Break loop
                            headIdx = UBound(readTagHeaders)
                        End If
                    Next headIdx
                End If
                        
            Case Constants.TASK_PRIORITY_HEADER
                'Task sheet does not have a priority value stored. Set to initial priority
                existingVal = Constants.TASK_PRIO_INITIAL
                    
            Case Else
                'Just pass the header here to retrieve the value of the sheet
                existingVal = Utils.GetSingleDataCellVal(sourceSheet, CStr(header))
        End Select
        
        'Now save the value to the new task and run handlers
        
        Dim cell As Range
        Set cell = PlanningUtils.IntersectHashAndListColumn(newHash, CStr(header))
        
        'Handle cell value changes here
        Select Case unifiedHeader
            Case TASK_NAME_HEADER
                'Copy name and handle name change to copy the name to the task sheet
                existingVal = existingVal + " (import)"
                cell.Value = existingVal
                Call Planning.HandleChanges(cell)
            
            Case TASK_ESTIMATE_HEADER
                cell.Value = existingVal
                Call Planning.ManageEstimateChange(cell, False)
                
            Case KANBAN_LIST_HEADER
                'Copy the value and manage change: Do not update the finished on date. This happens individually
                cell.Value = existingVal
                Call Planning.ManageKanbanListChange(cell, False)
                
            Case Constants.TAG_REGEX
                'Copy the tag and run handler without cell validation update. This causes cell validation to fail
                cell.Value = existingVal
                Call Planning.ManageTagChange(cell, False)
                
            Case Constants.CONTRIBUTOR_HEADER
                'Handle contributor change without cell validation update. This causes cell validation to fail
                'Also do not update the EBS estimates of the copied task
                cell.Value = existingVal
                Call Planning.ManageContributorChange(cell, False, False)
                                
            Case TASK_PRIORITY_HEADER, COMMENT_HEADER, DUE_DATE_HEADER, TASK_FINISHED_ON_HEADER
                'Copy values and handle change to copy values to task sheet
                cell.Value = existingVal
                Call Planning.HandleChanges(cell)
        End Select
    Next header
    
    'At the end sort the tasks with their priorities
    PlanningUtils.OrganizePrioColumn
End Function

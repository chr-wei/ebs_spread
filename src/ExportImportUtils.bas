Attribute VB_Name = "ExportImportUtils"
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

Const EXCHANGE_WORKBOOK_PREFIX As String = "EbsExportImport_"

Option Explicit



Function ExportVisibleTasks()
    Dim hashRange As Range
    
    Set hashRange = PlanningUtils.GetTaskListColumn(Constants.T_HASH_HEADER, ceData)
    Dim cll As Range
    Dim visibleHashes As Range
    For Each cll In hashRange
        If cll.Height <> 0 Then
            Set visibleHashes = Base.UnionN(visibleHashes, cll)
        End If
    Next cll
        
    'Debug info
    'Debug.Print "Visible hashes: " & visibleHashes.Count & " out of " & hashRange.Count
    
    If Not visibleHashes Is Nothing Then
        Dim exportWb As Workbook
        Set exportWb = Excel.Workbooks.Add
 
        
        Dim firstWorksheet As Worksheet
        Set firstWorksheet = exportWb.Worksheets(1)
        
        Dim sheet As Worksheet
                    
        For Each cll In visibleHashes
            Set sheet = TaskUtils.GetTaskSheet(cll.Value)
            
            Dim copiedSheet As Worksheet
            If Not sheet Is Nothing Then
                Call sheet.Copy(After:=firstWorksheet)
            End If
        Next cll
        
        For Each sheet In exportWb.Worksheets
            sheet.Visible = xlSheetVisible
        Next sheet
        
        Call Utils.DeleteWorksheetSilently(firstWorksheet)
    End If
End Function



Function ImportTasks(sourceWb As Workbook)
    
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
            
            Call PlanningUtils.CopyTask(newHash:=newHash, cpSource:=ceCopyFromWorksheetData, sourceSheet:=sheet)
        End If
    Next sheet
End Function

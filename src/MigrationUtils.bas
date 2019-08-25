Attribute VB_Name = "MigrationUtils"
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

Sub CopyEbsTableToTaskSheets()
'v0.97 and earlier to v0.98
'Replaces the old second table of task sheets with a simpler one
    Dim templateSheet As Worksheet
    Set templateSheet = Worksheets(Constants.TASK_SHEET_TEMPLATE_NAME)
    
    Dim copiedRange As Range
    Set copiedRange = templateSheet.Range("G13:K14")
    
    Dim taskSheet As Worksheet
    
    For Each taskSheet In Worksheets
        If SanityChecks.CheckHash(taskSheet.name) Then
            Dim deletedRange As Range
            Set deletedRange = taskSheet.Range("$G:$M")
            Call deletedRange.Delete
            
            Call copiedRange.Copy(taskSheet.Range("G13:K14"))
            taskSheet.Range("G13:K14").Columns.AutoFit
        End If
    Next taskSheet
End Sub



Sub ReplaceUserEstimateText()
'v0.97 and earlier to v0.98
    Dim sheet As Worksheet
    
    For Each sheet In Worksheets
        Dim found As Range
        Set found = Base.FindAll(sheet.UsedRange, "User estimate in h")
        If Not found Is Nothing Then found = "User time estimate"
    Next sheet
End Sub

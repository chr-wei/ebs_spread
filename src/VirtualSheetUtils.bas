Attribute VB_Name = "VirtualSheetUtils"
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

Const VIRTUAL_SHEET_NAME_HEADER  As String = "VIRTUAL_SHEET_NAME"
Const VIRTUAL_SHEET_STOR_ROWS_HEADER As String = "VIRTUAL_SHEET_RANGE_ROWS"
Const VIRTUAL_SHEET_STOR_COLS_HEADER As String = "VIRTUAL_SHEET_RANGE_COLS"
Const VIRTUAL_SHEET_OFFSET As Integer = 2



Sub Test_StoreAsVirtualSheet()
    Call StoreAsVirtualSheet(Worksheets("test_sheet"), Constants.STORAGE_SHEET_PREFIX)
    Call StoreAsVirtualSheet(Worksheets("full_sheet"), Constants.STORAGE_SHEET_PREFIX)
    Call StoreAsVirtualSheet(Worksheets("too_much"), Constants.STORAGE_SHEET_PREFIX)
End Sub



Sub Test_LoadVirtualSheet()
    Call LoadVirtualSheet("full_sheet", Constants.STORAGE_SHEET_PREFIX)
    Call LoadVirtualSheet("test_sheet", Constants.STORAGE_SHEET_PREFIX)
    Call LoadVirtualSheet("too_much", Constants.STORAGE_SHEET_PREFIX)
End Sub



Function StoreAsVirtualSheet(inSheet As Worksheet, storagePrefix As String, Optional deleteNonVirtualSheet As Boolean = True)
    'This function stores a real worksheet inside a storage sheet. The range in which the source sheet is stored is called a 'virtual' sheet
    '
    'Input args:
    '   inSheet:                The real worksheet that shall be stored
    '   storagePrefix:          A prefix to identify the virtual storage sheets
    '   deleteNonVirtualSheet:  Bool stating whether the stored sheet should be removed.
    '                           Duplicate data could be generated if this is set to true - be careful
    
    'Check args
    If inSheet Is Nothing Then Exit Function
    
    'Get data of the real sheet
    Dim rSheetRng As Range
    Set rSheetRng = inSheet.UsedRange
    
    If VirtualSheetUtils.VirtualSheetExists(inSheet.name, storagePrefix) Then
        'Only store if sheet is not already virtual (check by sheet name)
        Debug.Print "Cannot store sheet '" & inSheet.name & "' as virtual sheet. Sheet already exists.'"
        Exit Function
    End If
    
    'Prepare or get an existing virtual storage sheet
    Dim vSheetStorageRng As Range
    Set vSheetStorageRng = VirtualSheetUtils.GetNewSheetStorage(inSheet, storagePrefix)
       
    'Copy and paste the whole cell content of inSheet into the virtual storage range
    rSheetRng.Copy
    Call vSheetStorageRng.PasteSpecial(xlPasteAll)
    
    'Delete the inSheet after it has been stored inside the virtual inSheet
    If deleteNonVirtualSheet Then
        Call Utils.DeleteWorksheetSilently(inSheet)
    End If
End Function



Function GetFreeStorageSheet(inSheet As Worksheet, storagePrefix As String) As Worksheet
    'This function manages the virtual storage sheet(s) to store sheets in
    
    'Input args:
    '   inSheet:                The sheet that will be stored later. Test if enough storage exists.
    '   storagePrefix:          Prefix used to identify the storage sheet
    '
    'Output args:
    '   GetFreeStorageSheet:    The reference to a storage sheet which contains multiple virtual sheets
    
    'Init output
    Set GetFreeStorageSheet = Nothing
    
    'Check args
    If inSheet Is Nothing Then Exit Function

    Dim item As Variant
    Dim storageSheet As Worksheet
    For Each item In VirtualSheetUtils.GetAllStorageSheets(storagePrefix).Items
        Set storageSheet = item
        'Cycle through all available storage sheets
        If Not VirtualSheetUtils.StorageIsFull(storageSheet, inSheet) Then
            'Found a storage sheet with enough space
            Set GetFreeStorageSheet = storageSheet
            Exit Function
        End If
    Next item
    
    'No virtual storage available. Create new storage sheet
    Dim newStorage As Worksheet
    Set newStorage = ThisWorkbook.Worksheets.Add
    newStorage.name = Utils.CreateHashString(storagePrefix)
    Set GetFreeStorageSheet = newStorage
End Function



Function StorageIsFull(storageSheet As Worksheet, inSheet As Worksheet) As Boolean
    'Function checks if a storage sheet has enough space to store data in it
    '
    'Input args:
    '   storageSheet:   The storage sheet to test
    '   inSheet:        The sheet (data) that shall be made virtual. Defines the space to reserve
    '
    'Output args:
    '   StorageIsFull:  Boolean
    
    'Init output
    StorageIsFull = False
    
    'Check args
    If storageSheet Is Nothing Or inSheet Is Nothing Then Exit Function
    
    'Maximum rows that are available in the sheet (const)
    Dim maxRows As Long
    maxRows = storageSheet.Rows.Count
    
    If Utils.GetBottomLeftCell(storageSheet.UsedRange).Row + VIRTUAL_SHEET_OFFSET + inSheet.UsedRange.Rows.Count <= maxRows Then
        'Last row of used storage + offset + rows to store -> too many rows to store / not enough space available
        StorageIsFull = False
    Else
        StorageIsFull = True
    End If
End Function



Function LoadVirtualSheet(sheetName As String, storagePrefix As String, Optional templateSheet As Worksheet = Nothing) As Worksheet
    'The function loads a virtual sheet out of a storage sheet and inserts it into a (new) worksheet.
    '
    'Input args:
    '   sheetName:          The sheet's name you want to load
    '   storagePrefix:      The prefix to identify the storage sheet(s) containing virtual sheets
    '   templateSheet:      A template which is copied prior to copy the virtual sheet's data. A template can provide sheet code.
    '                       The virtual sheet only contains cell data
    '
    'Output args:
    '   LoadVirtualSheet:   Reference to the sheet that was loaded
    
    'Init output
    Set LoadVirtualSheet = Nothing
    
    'Check args
    If Utils.SheetExists(sheetName) Then
        Debug.Print "Virtual sheet '" & sheetName & "' will not be loaded. A non-virtual worksheet with the same name already exists.'"
        Exit Function
    End If
    
    'Check whether to use a template sheet or a new one
    Dim useSheetTemplate As Boolean: useSheetTemplate = False
    If Not templateSheet Is Nothing Then useSheetTemplate = True
    
    If Not VirtualSheetUtils.VirtualSheetExists(sheetName, storagePrefix) Then
        Debug.Print "Virtual sheet '" & sheetName & "' does not exist and cannot be loaded.'"
    Else
        'The virtual sheet exists and can be loaded without conflict
        Dim vr As Range
        Set vr = VirtualSheetUtils.GetVirtualStorageDataRange(sheetName, storagePrefix)
                
        'Get the non-virtual storage sheet (nvs)
        Dim nvs As Worksheet
        
        If useSheetTemplate Then
            'Copy template sheet as starting point or ..
            Call templateSheet.Copy(after:=templateSheet)
            Set nvs = ThisWorkbook.Worksheets(templateSheet.name & " (2)")
        Else
            '.. add a new sheet
            Set nvs = ThisWorkbook.Worksheets.Add
        End If
        
        'Set the name
        nvs.name = sheetName
        
        'Copy and paste data
        vr.Copy
        Call nvs.UsedRange.PasteSpecial(xlPasteAll)
        
        'Free virtual sheet storage
        Call VirtualSheetUtils.DeleteVirtualSheet(sheetName, storagePrefix)
        
        Set LoadVirtualSheet = nvs
    End If
End Function



Function DeleteVirtualSheet(sheetName As String, storagePrefix As String)
    'This function deletes a virtual sheet and runs a garbage collection afterwards to delete empty storage sheets
    
    'Input args:
    '   sheetName:      The virtual sheet's name you want to delete
    '   storagePrefix:  The virtual sheet storage prefix used to identify the storage sheet
    
    
    If Not VirtualSheetUtils.VirtualSheetExists(sheetName, storagePrefix) Then
        Debug.Print "Virtual sheet '" & sheetName & "' does not exist and cannot be deleted.'"
    Else
        'The virtual sheet exists and can be deleted
        Dim vr As Range
        Set vr = VirtualSheetUtils.GetVirtualStorageDataRange(sheetName, storagePrefix)
        
        'Free virtual sheet storage: Delete data, the header range and the footer range (empty row)
        Base.UnionN(Base.UnionN(vr.EntireRow, Utils.GetTopNeighbour(vr)).EntireRow, Utils.GetBottomNeighbour(vr).EntireRow).Delete
    End If
    
    'Run garbage collection to delete empty storage sheets
    Call VirtualSheetUtils.GarbageCollectStorageSheets(storagePrefix)
End Function



Function GarbageCollectStorageSheets(storagePrefix As String)
    'Function deletes storage sheet if they do not contain any virtual sheet data anymore
    '
    'Input args:
    '   storagePrefix:  Used to identify the storage sheet
    
    Dim item As Variant
    Dim storageSheet As Worksheet
    For Each item In VirtualSheetUtils.GetAllStorageSheets(storagePrefix).Items
        Set storageSheet = item
        If VirtualSheetUtils.IsStorageSheetEmpty(storageSheet) Then
            Call Utils.DeleteWorksheetSilently(storageSheet)
        End If
    Next item
End Function



Function GetVirtualStorageDataRange(sheetName As String, storagePrefix As String) As Range
    'Function returns the data range of a virtual sheet (a designated range inside a storage sheet)
    
    'Input args:
    '   sheetName: Name of the virtual sheet
    '   storagePrefix: The prefix used to identify the storage sheet
    '
    'Output args:
    '   GetVirtualStorageDataRange: The range the virtual sheet's data is in
    
    'Init output
    Set GetVirtualStorageDataRange = Nothing
    
    'Check args
    If StrComp(sheetName, "") = 0 Then Exit Function
    
    If VirtualSheetUtils.VirtualSheetExists(sheetName, storagePrefix) Then
        'Read the range as the virtual sheet exists: Read row and col count and span a range
        Dim vSheets As Scripting.Dictionary
        Set vSheets = VirtualSheetUtils.GetAllVirtualSheets(storagePrefix)
        
        Dim storageSheet As Worksheet
        Set storageSheet = vSheets(sheetName).Parent
        
        Dim nameCell As Range
        Set nameCell = vSheets(sheetName)
        
        'Read row and col count
        Dim rowCount As Long
        Dim colCount As Long
        rowCount = nameCell.Offset(0, 2).Value
        colCount = nameCell.Offset(0, 4).Value
        
        'Span the range
        Dim startRng As Range
        Dim endRng As Range
        
        Set startRng = nameCell.Offset(1, -1) 'Bottom left neighbour
        Set endRng = startRng.Offset(rowCount - 1, colCount - 1)
        Set GetVirtualStorageDataRange = storageSheet.Range(startRng, endRng)
        
        'Debug info
        'Debug.Print GetVirtualStorageDataRange.Address
    End If
End Function



Function GetNewSheetStorage(inSheet As Worksheet, storagePrefix As String) As Range
    'Return a range to store data in. The sheet storage will be prepared as follows:
    ' <VSHEET_1>
    ' <HEADER_ROW>: <NAME_HEADER> | <NAME> | <STORAGE_ROWS_HEADER> | <ROW_COUNT> | <STORAGE_COLUMNS_HEADER> | <COLUMN_COUNT>
    ' <DATA_AREA>:  <AREA WITH SIZE OF <ROW_COUNT> x <COLUMN_COUNT>> (This range will be returned
    ' <FOOTER_ROW>: A footer row (empty). Prevents auto list concat if list ends in the last row of the storage
    ' ...
    ' <VSHEET_N>
    '
    'Input args:
    '   inSheet:        The sheet you want to store and for which storage is reserved
    '   storagePrefix:  Prefix used to identify the storage sheet
    '
    'Output args:   The range the sheet can be stored in (<DATA_AREA>)
    
    'Get a storage sheet which has enough space to store the data of 'inSheet'
    Dim freeStorageSheet As Worksheet
    Set freeStorageSheet = VirtualSheetUtils.GetFreeStorageSheet(inSheet, storagePrefix)
    
    Dim usedRng As Range
    Set usedRng = freeStorageSheet.UsedRange
    
    'Get the bottom left cell of the existing data
    Dim vSheetNameHeaderCell As Range
    Set vSheetNameHeaderCell = Utils.GetBottomLeftCell(usedRng)
                
    If Not VirtualSheetUtils.IsStorageSheetEmpty(freeStorageSheet) Then
        'If the returned storage sheet already contains some storage sheets enter the header below an empty footer row
        Set vSheetNameHeaderCell = vSheetNameHeaderCell.Offset(VIRTUAL_SHEET_OFFSET, 0)
    End If
    
    'Set header and retrieve cell ranges
    
    'Save name
    vSheetNameHeaderCell.Value = VIRTUAL_SHEET_NAME_HEADER
    vSheetNameHeaderCell.Offset(0, 1).Value = inSheet.name
    
    'Save row count
    vSheetNameHeaderCell.Offset(0, 2).Value = VIRTUAL_SHEET_STOR_ROWS_HEADER
    vSheetNameHeaderCell.Offset(0, 3).Value = inSheet.UsedRange.Rows.Count
    
    'Save col count
    vSheetNameHeaderCell.Offset(0, 4).Value = VIRTUAL_SHEET_STOR_COLS_HEADER
    vSheetNameHeaderCell.Offset(0, 5).Value = inSheet.UsedRange.Columns.Count
    
    'Get the data storage
    Dim tlc As Range 'tlc = top left cell
    Dim brc As Range 'brc = bottom right cell
    
    'Span a range between top left and bottom right
    Set tlc = Utils.GetBottomNeighbour(vSheetNameHeaderCell)
    Set brc = tlc.Offset(inSheet.UsedRange.Rows.Count - 1, inSheet.UsedRange.Columns.Count - 1)
    
    Set GetNewSheetStorage = freeStorageSheet.Range(tlc, brc)
End Function



Function VirtualSheetExists(sheetName As String, storagePrefix As String) As Boolean
    'Test if a virtual sheet exists
    '
    'Input args:
    '   sheetName:         The virtual sheet's identifier
    '   storagePrefix:     The prefix the storage sheet is identified by
    '
    'Output args:
    '   VirtualSheetExists:  Boolean
    
    'Init output
    VirtualSheetExists = False
    
    'Check args
    If StrComp(sheetName, "") = 0 Then Exit Function
    
    'Load all virtual sheets to a collection of key=name, value=data range
    Dim vSheets As Scripting.Dictionary
    Set vSheets = VirtualSheetUtils.GetAllVirtualSheets(storagePrefix)

    VirtualSheetExists = vSheets.Exists(sheetName)
End Function



Function GetAllVirtualSheets(storagePrefix As String) As Scripting.Dictionary
    'Function returns a dictionary of all virtual sheets (key value pair of key:=sheet's name and value:= range of the found cell where the
    'name is stored
    '
    'Input args:
    '   storagePrefix:          The prefix used to identify the storage sheet containing the virtual sheet
    '
    'Output args:
    '   GetAllVirtualSheets:    A dictionary containing all virtual sheets of the workbook (key value pair, see above)
    
    'Init output
    Dim vSheets As New Scripting.Dictionary
    Set GetAllVirtualSheets = vSheets
    
    'Search for storage sheets
    Dim sheet As Worksheet
    Dim foundInSheet As Range
                
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.name Like storagePrefix & "*" Then
            
            'Search for virtual sheet entries inside the storage sheet. Do not use 'Base.FindAll' here, as it is much slower with many cells
            
            'Pass a column to search for virtual sheet name header const (first column of storage sheet)
            Dim sheetNameHeaderRange As Range
            Set sheetNameHeaderRange = Base.IntersectN(sheet.UsedRange, sheet.cells(1, 1).EntireColumn)
            
            'Get a range of all cells containing the virtual sheet name header const
            Set foundInSheet = VirtualSheetUtils.FindAllSheetNameHeaders(sheetNameHeaderRange)
            
            If Not foundInSheet Is Nothing Then
                'Concat all the sheet names
                Dim cll As Range
                Dim key As String
                For Each cll In foundInSheet
                    Dim nameCell As Range
                    Set nameCell = cll.Offset(0, 1) 'Get offset from name header (name val is stored right beside the header)
                    key = CStr(nameCell.Value)
                    
                    If Not vSheets.Exists(key) Then
                        'Store name cell with key
                        Call vSheets.Add(key:=key, item:=nameCell)
                    End If
                Next cll
            End If
        End If
    Next sheet
    
    Set GetAllVirtualSheets = vSheets
End Function



Function IsStorageSheetEmpty(sheet As Worksheet) As Boolean
    'Init output
    IsStorageSheetEmpty = True
    
    'Check args
    If sheet Is Nothing Then Exit Function
    If StrComp(sheet.UsedRange.Address, "$A$1") = 0 Then
        IsStorageSheetEmpty = True
    Else
        IsStorageSheetEmpty = False
    End If
End Function



Function FindAllSheetNameHeaders(ByVal rng As Range) As Range
    'Find a range of cells containing virtual sheet name headers
    
    'Input args:
    '  rng:                     The range in which searching is performed
    '
    'Output args:
    '  FindAllSheetNameHeaders: Range of cells matching the criteria (subset of rng)
    
    'Init output
    Set FindAllSheetNameHeaders = Nothing
    
    'Check args
    If rng Is Nothing Then Exit Function
    
    Dim cell As Range
    Dim result As Range
    
    Dim rngFirstMatch As Range
    Dim rngLastMatch As Range
    
    'Initial search to get a stop condition
    Set rngFirstMatch = rng.Find(VIRTUAL_SHEET_NAME_HEADER)
    Set rngLastMatch = rngFirstMatch
    
    'Find all other matches and set a new start cell in each value
    Dim strt As Range
    Dim rowOffset As Long
    
    Do
        'Cycle through the passed range and search for all virtual sheet name headers.
        'After every cycle start search with a row offset to gain speed (searched column range of the sheet can be large)
        
        Set result = Base.UnionN(result, rngLastMatch)
        'Do not search inside the data range of the virtual sheets (set row offset)
        rowOffset = rngLastMatch.Offset(0, 3).Value
        Set strt = rngLastMatch.Offset(rowOffset, 0)
        
        Set rngLastMatch = rng.FindNext(strt)
    Loop Until StrComp(rngFirstMatch.Address, rngLastMatch.Address) = 0 'Loop until the first match is found again
    
    Set FindAllSheetNameHeaders = result
    
    'Debug info
    'Debug.Print result.Address
End Function



Function GetAllStorageSheets(storagePrefix As String) As Scripting.Dictionary
    'Get all workbook virtual storage sheets matching the hash pattern
    '
    'Input args:
    '   The prefix used to identify the storage sheet
    '
    'Output args:
    '   GetAllStorageSheets:    Dictionary that contains the sheets in key value pair
    
    Dim storageSheets As New Scripting.Dictionary
    Dim sheet As Worksheet
    
    'Get task sheets (they have a hash set as their name)
    Dim sheetIdx: sheetIdx = 0
    For Each sheet In ThisWorkbook.Worksheets
        If VirtualSheetUtils.SheetIsStorageSheet(sheet, storagePrefix) Then
            Call storageSheets.Add(key:=sheet.name, item:=sheet)
        End If
    Next sheet

    Set GetAllStorageSheets = storageSheets
End Function



Function SheetIsStorageSheet(sheet As Worksheet, storagePrefix) As Boolean
    'Check if the sheet is a storage sheet (has special prefix)
    '
    'Input args:
    '   sheet:  The sheet to check
    '   storagePrefix: Prefix to identify the storage sheet
    '
    'Output args:
    '   SheetIsStorageSheet:    Boolean
    
    'Init output
    SheetIsStorageSheet = False
    
    'Check args
    If sheet Is Nothing Then Exit Function
    
    If sheet.name Like storagePrefix & "*" Then
        SheetIsStorageSheet = True
    Else
        SheetIsStorageSheet = False
    End If
End Function

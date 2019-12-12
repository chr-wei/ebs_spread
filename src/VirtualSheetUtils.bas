Attribute VB_Name = "VirtualSheetUtils"
Option Explicit

Const VIRTUAL_SHEET_NAME_HEADER  As String = "VIRTUAL_SHEET_NAME"
Const VIRTUAL_SHEET_STOR_ROWS_HEADER As String = "VIRTUAL_SHEET_RANGE_ROWS"
Const VIRTUAL_SHEET_STOR_COLS_HEADER As String = "VIRTUAL_SHEET_RANGE_COLS"
Const STORAGE_SHEET_PREFIX As String = "VSHEET_STOR_"



Sub Test_StoreVirtualSheet()
    Call StoreVirtualSheet(Worksheets("test_sheet"))
    Call StoreVirtualSheet(Worksheets("full_sheet"))
    Call StoreVirtualSheet(Worksheets("too_much"))
End Sub



Sub Test_LoadVirtualSheet()
    Call LoadVirtualSheet("full_sheet")
    Call LoadVirtualSheet("test_sheet")
    Call LoadVirtualSheet("too_much")
End Sub



Function StoreVirtualSheet(inSheet As Worksheet)

    'Check args
    If inSheet Is Nothing Then Exit Function
    
    Dim rSheetRng As Range
    Dim vSheetNameCell As Range
    Dim vSheetRowCountCell As Range
    Dim vSheetColCountCell As Range
    Dim vSheetStorageRng As Range
    
    Set rSheetRng = inSheet.UsedRange
    
    If VirtualSheetUtils.VirtualSheetExists(inSheet.name) Then
        Debug.Print "Cannot store sheet '" & inSheet.name & "' as virtual sheet. Sheet already exists.'"
        Exit Function
    End If
      
    Call VirtualSheetUtils.GetNewSheetStorage(inSheet, vSheetStorageRng)
       
    'Copy and paste the whole inSheet inside the virtual storage
    rSheetRng.Copy
    Call vSheetStorageRng.PasteSpecial(xlPasteAll)
    
    'Delete the inSheet after it has been stored inside the virtual inSheet
    Call Utils.DeleteWorksheetSilently(inSheet)
End Function



Function GetFreeStorageSheet(inSheet As Worksheet) As Worksheet
    'Init output
    Set GetFreeStorageSheet = Nothing
    
    'Check args
    If inSheet Is Nothing Then Exit Function

    Dim storageSheet As Worksheet
    
    For Each storageSheet In VirtualSheetUtils.GetAllStorageSheets
        If Not VirtualSheetUtils.StorageIsFull(storageSheet, inSheet) Then
            Set GetFreeStorageSheet = storageSheet
            Exit Function
        End If
    Next storageSheet
    
    'No virtual storage available. Create new storage sheet
    
    Dim newStorage As Worksheet
    Set newStorage = ThisWorkbook.Worksheets.Add
    newStorage.name = Utils.CreateHashString(STORAGE_SHEET_PREFIX)
    Set GetFreeStorageSheet = newStorage
End Function



Function StorageIsFull(storageSheet As Worksheet, inSheet As Worksheet) As Boolean
    'Init output
    StorageIsFull = False
    
    'Check args
    If storageSheet Is Nothing Or inSheet Is Nothing Then Exit Function
    
    Dim maxRows As Long
    maxRows = storageSheet.Rows.Count
    
    If Utils.GetBottomLeftCell(storageSheet.UsedRange).Row + 2 + inSheet.UsedRange.Rows.Count <= maxRows Then
        StorageIsFull = False
    Else
        StorageIsFull = True
    End If
End Function



Function LoadVirtualSheet(sheetName As String) As Worksheet
    
    'Init output
    Set LoadVirtualSheet = Nothing
    
    If Utils.SheetExists(sheetName) Then
        Debug.Print "Virtual sheet '" & sheetName & "' will not be loaded. A non-virtual worksheet with the same name already exists.'"
        Exit Function
    End If
    
    If Not VirtualSheetUtils.VirtualSheetExists(sheetName) Then
        Debug.Print "Virtual sheet '" & sheetName & "' does not exist and cannot be loaded.'"
    Else
        'The virtual sheet exists and can be loaded without conflict
        Dim vr As Range
        Set vr = VirtualSheetUtils.GetVirtualStorageDataRange(sheetName)
        
        vr.Copy
        
        Dim nvs As Worksheet
        
        Set nvs = ThisWorkbook.Worksheets.Add
        nvs.name = sheetName
        Call nvs.UsedRange.PasteSpecial(xlPasteAll)
        
        'Free virtual sheet storage
        Call VirtualSheetUtils.DeleteVirtualSheet(sheetName)
        
        Set LoadVirtualSheet = nvs
    End If
End Function



Function DeleteVirtualSheet(sheetName As String)
    
    If Not VirtualSheetUtils.VirtualSheetExists(sheetName) Then
        Debug.Print "Virtual sheet '" & sheetName & "' does not exist and cannot be deleted.'"
    Else
        'The virtual sheet exists and can be deleted
        Dim vr As Range
        Set vr = VirtualSheetUtils.GetVirtualStorageDataRange(sheetName)
        
        'Free virtual sheet storage
        Base.UnionN(Base.UnionN(vr.EntireRow, Utils.GetTopNeighbour(vr)).EntireRow, Utils.GetBottomNeighbour(vr).EntireRow).Delete
    End If
    
    Call VirtualSheetUtils.GarbageCollectStorageSheets
End Function



Function GarbageCollectStorageSheets()
    Dim storageSheet As Worksheet
    
    For Each storageSheet In VirtualSheetUtils.GetAllStorageSheets
        If VirtualSheetUtils.IsStorageSheetEmpty(storageSheet) Then
            Call Utils.DeleteWorksheetSilently(storageSheet)
        End If
    Next storageSheet
End Function



Function GetVirtualStorageDataRange(sheetName As String) As Range
    Set GetVirtualStorageDataRange = Nothing
    
    If VirtualSheetUtils.VirtualSheetExists(sheetName) Then
        Dim vSheets As Collection
        Set vSheets = VirtualSheetUtils.GetAllVirtualSheets
        
        Dim storageSheet As Worksheet
        Dim nameCell As Range
        
        Dim rowCount As Long
        Dim colCount As Long
        
        Set storageSheet = vSheets(sheetName).Parent
        Set nameCell = vSheets(sheetName)
        rowCount = nameCell.Offset(0, 2).Value
        colCount = nameCell.Offset(0, 4).Value
        
        Dim startRng As Range
        Dim endRng As Range
        
        Set startRng = Utils.GetLeftNeighbour(Utils.GetBottomNeighbour(nameCell))
        Set endRng = startRng.Offset(rowCount - 1, colCount - 1)
        Set GetVirtualStorageDataRange = storageSheet.Range(startRng, endRng)
        
        'Debug info
        'Debug.Print GetVirtualStorageDataRange.Address
    End If
    
End Function



Function GetNewSheetStorage(inSheet As Worksheet, _
    ByRef vSheetStorageRng As Range)
    
    'Get a storage sheet which has enough space to store the data of 'inSheet'
    Dim freeStorageSheet As Worksheet
    Set freeStorageSheet = VirtualSheetUtils.GetFreeStorageSheet(inSheet)
    
    Dim vSheetNameHeaderCell As Range
    Dim vSheetNameCell As Range
    Dim vSheetRowCountCell As Range
    Dim vSheetColCountCell As Range
    
    Dim usedRng As Range
    Set usedRng = freeStorageSheet.UsedRange
    
    Set vSheetNameHeaderCell = Utils.GetBottomLeftCell(usedRng)
                
    If Not VirtualSheetUtils.IsStorageSheetEmpty(freeStorageSheet) Then
        Set vSheetNameHeaderCell = Utils.GetBottomNeighbour(vSheetNameHeaderCell)
        Set vSheetNameHeaderCell = Utils.GetBottomNeighbour(vSheetNameHeaderCell)
    End If
    
    vSheetNameHeaderCell.Value = VIRTUAL_SHEET_NAME_HEADER
    Set vSheetNameCell = Utils.GetRightNeighbour(vSheetNameHeaderCell)
    vSheetNameCell.Value = inSheet.name
    
    'Set header and retrieve cell ranges for row count
    Dim vSheetRowCountHeaderCell As Range
    Set vSheetRowCountHeaderCell = Utils.GetRightNeighbour(vSheetNameCell)
    vSheetRowCountHeaderCell.Value = VIRTUAL_SHEET_STOR_ROWS_HEADER
    Set vSheetRowCountCell = Utils.GetRightNeighbour(vSheetRowCountHeaderCell)
    vSheetRowCountCell.Value = inSheet.UsedRange.Rows.Count
    
    'Set header and retrieve cell ranges for col count
    Dim vSheetColCountHeaderCell As Range
    Set vSheetColCountHeaderCell = Utils.GetRightNeighbour(vSheetRowCountCell)
    vSheetColCountHeaderCell.Value = VIRTUAL_SHEET_STOR_COLS_HEADER
    Set vSheetColCountCell = Utils.GetRightNeighbour(vSheetColCountHeaderCell)
    vSheetColCountCell.Value = inSheet.UsedRange.Columns.Count
    
    Dim tlc As Range
    Dim brc As Range
    
    Set tlc = Utils.GetBottomNeighbour(vSheetNameHeaderCell)
    Set brc = tlc.Offset(inSheet.UsedRange.Rows.Count - 1, inSheet.UsedRange.Columns.Count - 1)
    
    Set vSheetStorageRng = Range(tlc, brc)

End Function



Function VirtualSheetExists(sheetName As String) As Boolean

    'Init output
    VirtualSheetExists = False
    
    Dim vSheets As Collection
    Set vSheets = VirtualSheetUtils.GetAllVirtualSheets
    
    On Error Resume Next
    Call vSheets(sheetName)
    If Err.Number <> 0 Then
        VirtualSheetExists = False
    Else
        VirtualSheetExists = True
    End If
    
    On Error GoTo 0
End Function



Function GetAllVirtualSheets() As Collection
    Dim vSheets As New Collection
    
    'Init output
    Set GetAllVirtualSheets = vSheets
    
    Dim sheet As Worksheet
    Dim foundInSheet As Range
                
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.name Like STORAGE_SHEET_PREFIX & "*" Then
            
            'Search for virtual sheet entries inside the sheet. Do not use 'Base.FindAll' here, as it is much slower with many cells
            Dim sheetNameHeaderRange As Range
            Set sheetNameHeaderRange = Base.IntersectN(sheet.UsedRange, sheet.cells(1, 1).EntireColumn)
            
            Set foundInSheet = VirtualSheetUtils.FindAllSheetNameHeaders(sheetNameHeaderRange)
            
            If Not foundInSheet Is Nothing Then
                'Concat all the sheet names
                Dim cll As Range
                
                For Each cll In foundInSheet
                    Call vSheets.Add(cll.Offset(0, 1), CStr(cll.Offset(0, 1).Value))
                    Set vSheets = Base.GetUniqueStrings(vSheets)
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
    'xxFind a cell in a given range which matches a given value. By default a text comparison of the cell is performed.
    'This function also works for hidden cells
    
    'Input args:
    '  rng:           The range in which searching is performed
    '  propertyVal:    The value of the property one wants to find ('Value'property is default)
    '  compType:       The type of comparison one wants to use (<, >, etc.)
    '
    'Output args:
    '  FindAll:        Range of cells matching the criteria (subset of rng)

    
    'Init output
    Set FindAllSheetNameHeaders = Nothing
    
    'Check args
    If rng Is Nothing Then Exit Function
    
    Dim cell As Range
    Dim result As Range
    
    Dim rngFirstMatch As Range
    Dim rngLastMatch As Range
    
    Set rngFirstMatch = rng.Find(VIRTUAL_SHEET_NAME_HEADER)
    Set rngLastMatch = rngFirstMatch
    
    Dim strt As Range
    Do
        Set result = Base.UnionN(result, rngLastMatch)
        Dim rowOffset As Long
        rowOffset = rngLastMatch.Offset(0, 3).Value
        
        'On Error Resume Next
        Set strt = rngLastMatch.Offset(rowOffset, 0)
        
        'If Err.Number <> 0 Then
        '    Set strt = rngLastMatch.Offset(rowOffset, 0)
        'End If
        'On Error GoTo 0
        
        Set rngLastMatch = rng.FindNext(strt)
        
        'If Not rngLastMatch Is Nothing Then
            
        'End If
        
    Loop Until StrComp(rngFirstMatch.Address, rngLastMatch.Address) = 0
    
    Set FindAllSheetNameHeaders = result
    
    'Debug info
    'Debug.Print result.Address
End Function



Function GetAllStorageSheets() As Collection
    'Get all workbook virtual storage sheets matching the hash pattern
    
    Dim storageSheets As New Collection
    Dim sheet As Worksheet
    
    'Get task sheets (they have a hash set as their name)
    For Each sheet In ThisWorkbook.Worksheets
        If VirtualSheetUtils.SheetIsStorageSheet(sheet) Then
            Call storageSheets.Add(sheet)
        End If
    Next sheet

    Set GetAllStorageSheets = storageSheets
End Function



Function SheetIsStorageSheet(sheet As Worksheet) As Boolean
    'Init output
    SheetIsStorageSheet = False
    
    'Check args
    If sheet Is Nothing Then Exit Function
    
    If sheet.name Like STORAGE_SHEET_PREFIX & "*" Then
        SheetIsStorageSheet = True
    Else
        SheetIsStorageSheet = False
    End If
End Function

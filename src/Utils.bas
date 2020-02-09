Attribute VB_Name = "Utils"
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

Option Explicit

Public Enum ListRowSelect
    'This enum specifies values to select the table rows which should be fetched.
    ceHeader = 0
    ceData = 1
    ceAll = 2
End Enum

Public Enum DeserializedType
    ceString = 0
    ceDoubles = 1
    ceDates = 2
End Enum

Function CreateHashString(Optional prefix As String = "") As String
    'Function is used to create a unique hash value (first 18 chars of sha256) with prefix if needed.
    '
    'Input args:
    '   prefix: The prefix you want to add to the hash.
    
    Dim cl As clsSHA256
    
    Set cl = New clsSHA256
    'Call randomize function to get a 'new' hash
    Randomize
    CreateHashString = Left(prefix + UCase(cl.SHA256(CStr(Rnd()))), 18)
    Set cl = Nothing
    'Debug info
    Debug.Print "Created HASH is: " & CreateHashString
End Function



Function GetListColumn(sheet As Worksheet, listId As Variant, colIdentifier As Variant, rowIdentifier As ListRowSelect) As Range
    'Get the data of a column of a list / table.
    '
    'Input args:
    '  sheet:  The sheet one wants to retrieve a list with a specific column for
    '  listId: The id (name or index) of the list you want to retrieve
    '  colIdentifier:  The cell of a column or its header name
    '  rowIdentifier:  A selector which returns either the header section, the data body range section or header and body section of the selected column.
    '
    'Output args:
    '  GetListColumn: The range of the column (header, body or body and header)
    
    'Init output
    Set GetListColumn = Nothing
    
    'Check args
    If sheet Is Nothing Then
        Exit Function
    End If
    
    Dim lo As ListObject
    
    'List identifier can either be string or integer
    Select Case (TypeName(listId))
        Case "Integer", "String"
            Set lo = sheet.ListObjects(listId)
    End Select
    
    'First find the list column
    Dim lc As ListColumn
    Select Case (TypeName(colIdentifier))
        Case "Integer"
            Set lc = lo.ListColumns(CInt(colIdentifier))
        Case "String"
            If Base.FindAll(lo.HeaderRowRange, colIdentifier) Is Nothing Then
                Exit Function
            End If
            Set lc = lo.ListColumns(colIdentifier)
        Case "Range"
            'Find list column via cell inside the column. Read the header name
            Dim cell As Range
            Dim header As String
            Set cell = colIdentifier
            header = Utils.GetListColumnHeader(cell)
            Set lc = lo.ListColumns(header)
    End Select
    
    'Then select all the rows of that column that should be returned
    Select Case rowIdentifier
        Case ListRowSelect.ceHeader
            'Intersect the header row and the column range
            Set GetListColumn = Base.IntersectN(lo.HeaderRowRange, lc.Range)
        Case ListRowSelect.ceData
            Set GetListColumn = lc.DataBodyRange
        Case ListRowSelect.ceAll
            Set GetListColumn = lc.Range
    End Select
End Function



Function GetListColumnHeader(cells As Range) As String
    'Source: https://stackoverflow.com/questions/47443844/find-column-name-of-active-cell-within-listobject-in-excel
    
    'Function to get the header of a list's column if a cell of that column is passed.
    '
    'Input args:
    '  cells:                   Range of cells of a column. Function returns no header if cells spanning multiple columns are passed
    '
    'Output args:
    '  GetListColumnHeader:     The header of the column
    
    'Init output
    GetListColumnHeader = ""
            
    If cells.Columns.Count <> 1 Then
        Exit Function
    End If
    
    If cells.ListObject Is Nothing Then
        Exit Function
    End If
    
    'Intersect header row and cell column to get the header
    Dim headerCell As Range
    Set headerCell = Base.IntersectN(cells.ListObject.HeaderRowRange, cells.EntireColumn)
    
    If Not headerCell Is Nothing Then
        GetListColumnHeader = headerCell.Value
    End If
    
    'Debug info
    'Debug.Print cells.Address
End Function



Function FindSheetCell(sheet As Worksheet, findString As String, Optional compType As ComparisonTypes = ComparisonTypes.ceStringComp) As Range
    'Wrapper function to find all cells of a given worksheet which values match the passed string
    '
    'Input args:
    '  sheet: The sheet one wants to find a value in
    '  findSring: The value one wants to find
    '  compType: The comparison mehtod used to compare the cell value (vba 'Like'method, regexp or double comparison
    
    'Init output
    Set FindSheetCell = Nothing
    
    Dim foundCells As Range
    Dim sheetRange As Range
    If StrComp(Trim(findString), "") <> 0 Then
        'Only search if a value was passed. Only search in the used range
        Set sheetRange = sheet.UsedRange
        Set foundCells = Base.FindAll(sheetRange, findString, , compType)
    End If
    
    Set FindSheetCell = foundCells
End Function



Function GetTopLeftCell(rng As Range) As Range
    'Function returns the top left cell of a given range. This method was written to prevent errors with the cells.End(xlDown) property.
    '
    'Input args:
    '   rng:            Range of any size to get the top left cell of
    '
    'Output args:
    '   GetTopLeftCell: The top left cell of the range
    
    'Init output
    Set GetTopLeftCell = Nothing
    
    'Check args
    If rng Is Nothing Then Exit Function
    Set GetTopLeftCell = rng.cells(1, 1)
End Function



Function GetBottomRightCell(rng As Range) As Range
    'Function returns the bottom right cell of a given range. This method was written to prevent errors with the cells.End(xlDown) property.
    '
    'Input args:
    '   rng:                Range of any size to get the top left cell of
    '
    'Output args:
    '   GetBottomRightCell: The bottom right cell of the range
    
    'Init output
    Set GetBottomRightCell = Nothing
    'Check args
    If rng Is Nothing Then Exit Function
    Set GetBottomRightCell = rng.cells(rng.Rows.Count, rng.Columns.Count)
End Function



Function GetBottomLeftCell(rng As Range) As Range
    'Function returns the bottom left cell of a given range. This method was written to prevent errors with the cells.End(xlDown) property.
    '
    'Input args:
    '   rng:                Range of any size to get the top left cell of
    '
    'Output args:
    '   GetBottomLeftCell: The bottom right cell of the range
    
    'Init output
    Set GetBottomLeftCell = Nothing
    'Check args
    If rng Is Nothing Then Exit Function
    Set GetBottomLeftCell = rng.cells(rng.Rows.Count, 1)
End Function



Function GetLeftNeighbour(rng As Range) As Range
    'Get a single next left cell to a given selection
    Set GetLeftNeighbour = Utils.GetTopLeftCell(rng).Offset(0, -1)
End Function



Function GetRightNeighbour(rng As Range) As Range
    'Get a single next right cell to a given selection
    Set GetRightNeighbour = Utils.GetBottomRightCell(rng).Offset(0, 1)
End Function



Function GetBottomNeighbour(rng As Range) As Range
    'Get a single next bottom cell to a given selection
    Set GetBottomNeighbour = Utils.GetBottomRightCell(rng).Offset(1, 0)
    
    'Debug info
    'Debug.Print (GetBottomNeighbour.Address)
End Function



Function GetTopNeighbour(rng As Range) As Range
    'Get a single next top cell to a given selection
    Set GetTopNeighbour = Utils.GetTopLeftCell(rng).Offset(-1, 0)
End Function



Function GetNumberedEntries(sheet As Worksheet, listId As Variant) As Range
    'Returns the range of all valid entries in a list. Every valid entry has a #00001 number in the first column i.e. it is not empty
    '
    'Input args:
    '   sheet:  The sheet which contains the list
    '   listId: The id of the list which entries should be examined
    '
    'Output args:
    '   GetNumberedEntries: The range of valid entries that are contained in the list
    
    'Init output
    Set GetNumberedEntries = Nothing
    
    'Read the entries of the first column of a list
    Dim entries As Range
    Set entries = Utils.GetListColumn(sheet, listId, 1, ceData)
    If entries Is Nothing Then
        Set GetNumberedEntries = Nothing
        Exit Function
    End If
    
    If entries.Rows.Count = 1 And StrComp(entries(1).Value, "") = 0 Then
        'The list always has one body row even if it is empty. If no entry is in the first row (empty string) then set the body range to nothing.
        Set GetNumberedEntries = Nothing
    Else
        Set GetNumberedEntries = entries
    End If
    'Debug info
    'Debug.Print "Found numbered rows in: " & FindNumberedEntries.Address
End Function



Function GetNextEntryNumber(numberedEntries As Range) As Long
    'Read the entry numbers of the passed entries and sort them to get the next valid entry number.
    'Numbers have '#00001' like format so extraction is a bit more complicated
    '
    'Input args:
    '   numberedEntries:    The range of a list which contains the numbered entries
    '
    'Output args:
    '   GetNextEntryNumber: Long containing the next entry number
    
    'Init output
    GetNextEntryNumber = -1
    
    'Check args
    If numberedEntries Is Nothing Then
        GetNextEntryNumber = 1
        Exit Function
    End If
    
    If numberedEntries.Count = 0 Then
        GetNextEntryNumber = 1
        Exit Function
    End If
    
    Dim entryNumbers() As Long
    Dim entryCount As Long
    Dim cell As Range
    Dim iter As Long
    
    entryCount = numberedEntries.Count
    ReDim entryNumbers(0 To entryCount - 1)
    iter = 0
    
    For Each cell In numberedEntries
        'Add the entry numbers to the array
        Dim currentNumberedString As String
        Dim croppedString As String
        Dim currentNumber As Long
        currentNumberedString = cell.Value

        Dim regex As New RegExp

        regex.Global = True
        regex.Pattern = "#"

        croppedString = regex.Replace(currentNumberedString, "")
        currentNumber = val(croppedString)
        entryNumbers(iter) = currentNumber
        iter = iter + 1
    Next cell
    
    'Sort the array
    Call QuickSort(entryNumbers, ceAscending)
    
    GetNextEntryNumber = entryNumbers(UBound(entryNumbers)) + 1
End Function



Function GetNewEntry(sheet As Worksheet, listId As Variant, ByRef newEntry As Range, ByRef newFormattedNumber As String) As Boolean
    'Returns a new valid entry (line) of a list to insert data in.
    '
    'Input args:
    '   sheet:  The sheet containing the list you want to add an entry in
    '   listId: The id of the list the entry should be added to. A number will be added to the first column of the list
    '
    'Output args:
    '   newEntry:   A cell containing the new number. Use it as reference to the current data row
    '   newFormattedNumber:  Number with format like #000001. The newest entry will get the highest number
    
    'Init output
    GetNewEntry = False
    
    'Check args
    If sheet Is Nothing Or StrComp(CStr(listId), "") = 0 Then
        Exit Function
    End If
    
    Dim entries As Range
    Dim head As Range
    Set entries = Utils.GetNumberedEntries(sheet, listId)
    
    If entries Is Nothing Then
        'If no entries could be found get the header cell. Next cell will be selected below the header cell. GetNumberedEntries will return 'Nothing'
        'if there is a data row in the list which has no numbered entry in the first column (means no entry for us)
        Set head = Utils.GetListColumn(sheet, listId, 1, ceHeader)
        
        'Debug info
        'Debug.Print entries.Address
    Else
        'Add a row to the list
        Dim lo As ListObject
        Set lo = sheet.ListObjects(listId)
        Call lo.ListRows.Add
    End If
    
    'Format the new number
    newFormattedNumber = "#" & Format(Utils.GetNextEntryNumber(entries), Constants.COUNTED_ENTRIES_FORMAT)
    
    'Append a new entry to the existing ones. If entries was nothing head is set, if entries was something head is not set
    Set newEntry = Utils.GetBottomNeighbour(Base.UnionN(head, entries))
    GetNewEntry = True
End Function



Function IntersectListColAndCells(sheet As Worksheet, listId As Variant, colIdentifier As Variant, rowIdentifier As Range) As Range
    'Intersect a range of rows with a single column of a list object
    'The function is mostly used to get single cells from a list object
    '
    'Input args:
    '   sheet:          The sheet the list object resides on
    '   listId:         The id of the list. Can be 'String' type or 'Integer' type
    '   colIdentifier:  The column you want to intersect. Can be a header 'String' or a cell 'Range' in the column.
    '   rowIdentifier:  A range of cells marking rows inside the list object
    
    'Init function
    Set IntersectListColAndCells = Nothing
    
    'Check args
    If (sheet Is Nothing Or StrComp(colIdentifier, "") = 0 Or rowIdentifier Is Nothing) Then
        Exit Function
    End If
    
    Dim listCol As Range

    Set listCol = Utils.GetListColumn(sheet, listId, colIdentifier, ceAll)
    
    If Not listCol Is Nothing Then
        Set IntersectListColAndCells = _
            Base.IntersectN(listCol, rowIdentifier.EntireRow)
    End If
    'Debug info
End Function



Function SetCellValuesForValidation(cellRange As Range)
    'This function updates all tag fields. It scans a range of cells for their string values and puts all unique strings into a data validation
    'filter to filter the fields themself. As a result new values can be set via drop down menu.
    
    'Input args:
    '   cellRange:  The range of cells you want to set the validation for
    
    Dim cell As Range
    
    'All strings in the passed cells. These are filtered to get unique values.
    Dim stringList As Collection
    Set stringList = Utils.ConvertRngToStrCollection(cellRange)
    
    Set stringList = Base.GetUniqueStrings(stringList)
    If stringList.Count > 0 Then
        
        'Build the list string
        Dim listString As String
        listString = Join(Base.CollectionToArray(stringList), ",")
        'Debug.Print "ls: " & listString
        
        Dim currArea As Variant
        Dim currCell As Variant
        
        'Set up the range validation (title, error msg etc.)
        With cellRange.Validation
            Call .Delete
            Call .Add(xlValidateList, xlValidAlertInformation, xlBetween, listString, "")
            .InputTitle = "List value"
            .ErrorTitle = "List value"
            .InputMessage = "Select or enter a new item"
            .ErrorMessage = "Enter a new entry inside this column?"
        End With
    End If
End Function



Function RunTryCatchedCall(f As String, Optional obj As Object, _
    Optional arg1 As Variant, Optional arg2 As Variant, Optional arg3 As Variant, Optional arg4 As Variant, _
    Optional enableEvt As Boolean = False, Optional screenUpdating As Boolean = True)
    'This is a wrapper to run functions of objects (worksheets) and modules with specific Application settings e.g. disabled Application events
    'with a try catch mechanism to prohibit errors from bothering the user
    
    'Comment next line to enable properly debugging
    If Not Constants.DEBUGGING_MODE Then
        On Error GoTo errHandle
    End If
    'Deactivate events and screen updating if necessary
    If Not enableEvt Then
        Excel.Application.EnableEvents = False
    End If
    If Not screenUpdating Then
        Excel.Application.screenUpdating = False
    End If
    
    'Call the passed function either on a passed object (e.g. worksheet) or just without (e.g a module function)
    If IsMissing(obj) Or obj Is Nothing Then
        Call Excel.Application.Run(f, arg1, arg2, arg3, arg4)
    Else
        If IsMissing(arg1) Then
            Call CallByName(obj, f, VbMethod)
        ElseIf IsMissing(arg2) Then
            Call CallByName(obj, f, VbMethod, arg1)
        ElseIf IsMissing(arg3) Then
            Call CallByName(obj, f, VbMethod, arg1, arg2)
        ElseIf IsMissing(arg4) Then
            Call CallByName(obj, f, VbMethod, arg1, arg2, arg3)
        Else
            Call CallByName(obj, f, VbMethod, arg1, arg2, arg3, arg4)
        End If
    End If
    
    GoTo finallyHandle 'Skip the error handle
    
errHandle:
    With Err
        If .Number <> 0 Then 'ein Fehler ist aufgetreten
            Select Case .Number
                Case Else
                    'Fehler-Meldung anzeigen und Prozedur beenden
                    MsgBox "Error no. " & .Number & vbLf & .Description
            End Select
        End If
    End With
        
finallyHandle:
    'Finally all other tasks
    Excel.Application.EnableEvents = True
    Excel.Application.screenUpdating = True
End Function



Function ConvertRangeValsToArr(rng As Range) As String()
    'Function gets all values of a range and puts them into an array
    '
    'Input args:
    '   rng:                        The range of cells you want to get the values of
    '
    'Output args:
    '   ConvertRangeValsToArr:   The array with values
    
    'Check args
    If rng Is Nothing Then Exit Function
    
    Dim rngCell As Range
    Dim rngValues() As String
    ReDim Preserve rngValues(rng.Count - 1)

    Dim rngIdx As Long
    rngIdx = 0

    For Each rngCell In rng
        rngValues(rngIdx) = rngCell.Value2
        rngIdx = rngIdx + 1
    Next rngCell

    ConvertRangeValsToArr = rngValues
End Function



Function SerializeArray(arr As Variant) As String
    'Function to write array values to a single cell after putting them into a serialized string
    '
    'Input args:
    '   arr:            The array you want to serialize
    '
    'Output args:
    '   SerializeArray: The serialized array string
    
    'Init output
    SerializeArray = ""
    
    'Check args
    If Not Base.IsArrayAllocated(arr) Then Exit Function
    
    Dim item As Variant
    
    'Create an helper array with equal size as the input array
    Dim strArr() As String
    ReDim strArr(UBound(arr))
    
    Dim idx As Long
    idx = 0
    For Each item In arr
        If StrComp(TypeName(item), "Double") = 0 Then
            'Write double with 4-digit precision to reduce storage space
            strArr(idx) = Format(item, "0.000")
        Else
            strArr(idx) = CStr(item)
        End If
    idx = idx + 1
    Next item
    
    'Build the string
    SerializeArray = "{" + Join(strArr, ";") + "}"
End Function



Function DeserializeArray(serialized As String) As Variant
    'Function to read a serialized array into an array.
    '
    'Input args:
    '   serialized: The string that contains the serialized data
    '
    'Output args:
    '   DeserializeArray:   The array of deserialized strings
    
    'Init output
    Dim deserialized As Variant
    DeserializeArray = deserialized
    
    'Check args
    Dim regex As New RegExp

    regex.Global = True
    regex.Pattern = Constants.SERIALIZED_ARRAY_REGEX

    If Not regex.test(serialized) Then
        Exit Function
    End If
    
    'Begin deserialization
    
    'Remove the braces embracing the string and split the string by delimiter
    'Deserializing strings with braces as content will result in errors
    
    serialized = Replace(serialized, "{", "")
    serialized = Replace(serialized, "}", "")
    deserialized = Split(serialized, ";")
    
    DeserializeArray = deserialized
End Function



Function ConvertRngToStrCollection(rng As Range) As Collection
    'Convert any range values to a collection of strings
    '
    'Input args:
    '   rng:        The range you want to convert
    '
    'Output args:
    '   ConvertRngToStrCollection:  The collection of strings
    
    Dim stringCollection As New Collection
    Set ConvertRngToStrCollection = stringCollection

    If rng Is Nothing Then
        Exit Function
    End If

    Dim cell As Range
    For Each cell In rng
        Dim currentValue As String
        currentValue = cell.Value
        ''Normal' add here without 'key'
        Call stringCollection.Add(currentValue)
    Next cell

    Set ConvertRngToStrCollection = stringCollection
End Function



Function SheetExists(name As String) As Boolean
    'Check if the sheet already exists.
    '
    'Input args:
    '   name:   The name you want to test
    '
    'Output args:
    '   SheetExists:    True if the sheet exists in the workbook
    
    'Check args
    If StrComp(name, "") = 0 Then Exit Function
    'Init output
    SheetExists = False
    
    Dim sheet As Worksheet
    
    'Cycle through all names and find the sheet
    For Each sheet In ThisWorkbook.Worksheets
        If StrComp(sheet.name, name) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next sheet
End Function



Function InterpolateArray(arrX() As Double, arrY() As Double, supportPoints As Variant) As Double()
    'For sorted x values with given y values an interpolation for a support point is performed
    '
    'Input args:
    '   arrX:               The array containing the x data
    '   arrY:               Array containing the y-values to the given x-values
    '   supportPoints:      Numeric array or single value containing 'support' points (x-axis dimension) for which one wants to interpolated
    '                       y-values for
    '
    'Output args:
    '   InterpolateArray:   The interpolated values for n support points
    
    'Init output
    Dim interpolatedY() As Double
    InterpolateArray = interpolatedY
    
    'Check args
    If UBound(arrX) < 1 Or UBound(arrX) <> UBound(arrY) Then
        'No interpolation with only one data point possible or inequal array lengths possible
        Exit Function
    End If
    
    Dim supportIsArray As Boolean
    supportIsArray = IsArray(supportPoints)
    
    If supportIsArray And Not Base.IsArrayAllocated(supportPoints) Then
        Exit Function
    End If
    
    'Copy the arrays to prevent modification of the calling function
    Dim copiedArrX() As Double
    copiedArrX = arrX
    Dim copiedArrY() As Double
    copiedArrY = arrY

    Call Base.QuickSort(copiedArrX, ceAscending, , , copiedArrY)
        
    'Data prepared now start calculating the interpolated values
    
    Dim searchIdx As Long
    Dim xSpan As Double
    Dim ySpan As Double
    Dim xRef As Double
    Dim yRef As Double
    Dim currentSupPoint As Double
    
    Dim interpolateIdx As Integer
    Dim interpolateLimit
    
    If supportIsArray Then
        ReDim interpolatedY(UBound(supportPoints))
        interpolateLimit = UBound(supportPoints)
    Else
        'Only one value has to be calculated - init array for only one value
        ReDim interpolatedY(0)
        interpolateLimit = 0
    End If
    
    'Iterate over all support points for which one whishes to interpolate
    For interpolateIdx = 0 To interpolateLimit
        If supportIsArray Then
            currentSupPoint = supportPoints(interpolateIdx)
        Else
            currentSupPoint = supportPoints
        End If
        
        searchIdx = LBound(copiedArrX)
        
        'Search for the inserted support points
        While copiedArrX(searchIdx) < currentSupPoint And searchIdx < UBound(copiedArrX)
            searchIdx = searchIdx + 1
        Wend
        
        If searchIdx = LBound(copiedArrX) Then
            xSpan = copiedArrX(LBound(copiedArrX) + 1) - copiedArrX(LBound(copiedArrX))
            ySpan = copiedArrY(LBound(copiedArrY) + 1) - copiedArrY(LBound(copiedArrY))
            xRef = copiedArrX(LBound(copiedArrX))
            yRef = copiedArrY(LBound(copiedArrY))
            
        ElseIf searchIdx = UBound(copiedArrX) Then
            xSpan = copiedArrX(UBound(copiedArrX)) - copiedArrX(UBound(copiedArrX) - 1)
            ySpan = copiedArrY(UBound(copiedArrY)) - copiedArrY(UBound(copiedArrY) - 1)
            xRef = copiedArrX(UBound(copiedArrX))
            yRef = copiedArrY(UBound(copiedArrY))
        Else
            xSpan = copiedArrX(searchIdx + 1) - copiedArrX(searchIdx - 1)
            ySpan = copiedArrY(searchIdx + 1) - copiedArrY(searchIdx - 1)
            xRef = copiedArrX(searchIdx - 1)
            yRef = copiedArrY(searchIdx - 1)
        End If
        Dim cal As Double
        cal = (currentSupPoint - xRef) / xSpan * ySpan + yRef
        
        If supportIsArray Then
            interpolatedY(interpolateIdx) = cal
        Else
            interpolatedY(0) = cal
        End If
    Next interpolateIdx
    
    InterpolateArray = interpolatedY
End Function



Function CopyVarArrToStringArr(varArr As Variant) As String()
    'Conversion method to get string array from (unknown type) variant array. See other conversion methods as well
    '
    'Input args:
    '   varArr:                 The variant array to convert
    '
    'Output args:
    '   CopyVarArrToStringArr:  The string typed array
    
    'Init output
    Dim stringArr() As String
    CopyVarArrToStringArr = stringArr
    
    'Check args
    If Not Base.IsArrayAllocated(varArr) Then Exit Function
    ReDim stringArr(UBound(varArr) - LBound(varArr))
    
    Dim varVal As Variant
    Dim elemIdx As Integer
    
    For elemIdx = 0 To UBound(varArr) - LBound(varArr)
        If GetArrayDimension(varArr) > 1 Then
            'If you read value of multiple cells the arrays have more than one dimension - catch that here
            stringArr(elemIdx) = CStr(varArr(elemIdx + 1, LBound(varArr)))
        Else
            stringArr(elemIdx) = CStr(varArr(elemIdx))
        End If
    Next elemIdx
    
    CopyVarArrToStringArr = stringArr
End Function



Function CopyVarArrToDoubleArr(varArr As Variant) As Double()
    'Conversion method to get double array from (unknown type) variant array. See other conversion methods as well
    '
    'Input args:
    '   varArr:                 The variant array to convert
    '
    'Output args:
    '   CopyVarArrTodoubleArr:  The double typed array
    
    'Init output
    Dim doubleArr() As Double
    CopyVarArrToDoubleArr = doubleArr
    
    'Check args
    If Not Base.IsArrayAllocated(varArr) Then Exit Function
    ReDim doubleArr(UBound(varArr) - LBound(varArr))
    
    Dim varVal As Variant
    Dim elemIdx As Integer
    
    For elemIdx = 0 To UBound(varArr) - LBound(varArr)
        If GetArrayDimension(varArr) > 1 Then
            'If you read value of multiple cells the arrays have more than one dimension - catch that here
            varVal = varArr(elemIdx + 1, LBound(varArr))
        Else
            varVal = varArr(elemIdx)
        End If
        
        'Check double conversion
        If IsNumeric(varVal) Or IsDate(varVal) Then
            'Works for Strings and numbers
            doubleArr(elemIdx) = varVal
        Else
            Erase doubleArr
            Exit Function
        End If
    Next elemIdx

    CopyVarArrToDoubleArr = doubleArr
End Function



Function CopyVarArrToDateArr(varArr As Variant) As Date()
    'Conversion method to get date array from (unknown type) variant array. See other conversion methods as well
    '
    'Input args:
    '   varArr:                 The variant array to convert
    '
    'Output args:
    '   CopyVarArrToDateArr:  The date typed array
    
    'Init output
    Dim dateArr() As Date
    CopyVarArrToDateArr = dateArr
    
    'Check args
    If Not Base.IsArrayAllocated(varArr) Then Exit Function
    ReDim dateArr(UBound(varArr) - LBound(varArr))
    
    Dim varVal As Variant
    Dim elemIdx As Integer
    
    For elemIdx = 0 To UBound(varArr) - LBound(varArr)
        
        If Base.GetArrayDimension(varArr) > 1 Then
            'If you read value of multiple cells the arrays have more than one dimension - catch that here
            varVal = varArr(elemIdx + 1, LBound(varArr))
        Else
            varVal = varArr(elemIdx)
        End If
        
        'Check date conversion
        If IsNumeric(varVal) Or IsDate(varVal) Then
            'Works for Strings and numbers
            dateArr(elemIdx) = varVal
        Else
            Erase dateArr
            Exit Function
        End If
    Next elemIdx
    
    CopyVarArrToDateArr = dateArr
End Function



Function GetSingleDataCellVal(sheet As Worksheet, headerText As String, Optional ByRef dataCell As Range) As String
    'Read a value of a cell that has a 'header' specifier to its left. The cell is identified by that header (cell)
    '
    'Input args:
    '   sheet:  The sheet the value is to be found
    '   headerText: The header text to search
    '
    'Output args:
    '   dataCell:               The function also returns the cell which holds the value
    '   GetSingleDataCellVal:   String containing the value right to the header cell
    
    'Init output
    GetSingleDataCellVal = ""
    
    Dim headerCell As Range
    Set headerCell = Utils.FindSheetCell(sheet, headerText)
    
    If Not headerCell Is Nothing Then
        Set headerCell = Utils.GetTopLeftCell(headerCell)
        Set dataCell = Utils.GetRightNeighbour(headerCell)
        
        If Not IsError(dataCell.Value) Then
            GetSingleDataCellVal = dataCell.Value
        End If
    End If
End Function



Function SetSingleDataCell(sheet As Worksheet, headerText As String, data As Variant)
    'Set a value of a cell that has a 'header' specifier to its left. The cell is identified by that header (cell)
    '
    'Input args:
    '   sheet:                  The sheet the header is to be found
    '   headerText:             The header text to search
    '   data:                   The data you want to set to the cell

    Dim headerCell As Range
    Dim dataCell As Range
    
    Set headerCell = Utils.FindSheetCell(sheet, headerText)
    If Not headerCell Is Nothing Then
        Set dataCell = Utils.GetRightNeighbour(headerCell)
        dataCell.Value = data
    Else
        'Debug info
        Debug.Print "Header cell '" & headerText & " 'could not be found on worksheet '" & sheet.name & "'."
    End If
End Function



Function AddSubtileHyperlink(cell As Range, link As String)
    'Add a hyperlink that has no underline and a subtile color
    '
    'Input args:
    '   cell:   The cell the hyperlink will be added to
    '   link:   The hyperlink to add
    
    With cell.Parent.Hyperlinks.Add(cell, "", link)
        Dim lightColor As Long
        lightColor = SettingUtils.GetColors(ceLightColor)
        .Range.Font.color = lightColor
        .Range.Font.Underline = False
    End With
End Function



Function DeleteFilteredListObjectRow(rowIdentifier As Range)
    'Function to delete a row of a list object identified by a cell of that row. Deletion is a bit tricky because the row number cannot be read
    'directly. In addition no rows can be deleted when filtering is active. Disable and reenable filters to delete the row.
    '
    'Input args:
    '   rowIdentifier:  A cell or multiple cells of one row identifying the row

    'Check args:
    If rowIdentifier Is Nothing Then Exit Function

    If rowIdentifier.Rows.Count > 1 Then
        Exit Function
    End If

    Dim lo As ListObject
    Set lo = rowIdentifier.ListObject

    If lo Is Nothing Then Exit Function

    If Not Base.IntersectN(rowIdentifier, lo.HeaderRowRange) Is Nothing Then
        'You selected the header row. Deleting the header row is not allowed
        Exit Function
    End If

    If Not lo Is Nothing Then
        Dim sheetRow As Long
        Dim headerRow As Long
        Dim rowIdx As Long
        sheetRow = rowIdentifier.Row
        headerRow = lo.HeaderRowRange.Row

        'Get the row idx with respect to the list object
        rowIdx = sheetRow - headerRow

        'Unhide all data as deletion only works when all cells are visible:
        'Read the active filters and store their props
        Dim filterData() As Variant
        filterData = Utils.GetAutoFilters(lo)
        
        'Display all data (filters are deleted)
        lo.AutoFilter.ShowAllData
        
        'Debug info
        'lo.ListRows(rowIdx).Range.Select
        'Delete the data row
        lo.ListRows(rowIdx).Delete
        
        'Reapply the filters again if there were any
        If Base.IsArrayAllocated(filterData) Then
            Call Utils.ReapplyAutoFilters(lo, filterData)
        End If
    End If
End Function



Function GetAutoFilters(lo As ListObject) As Variant()
    'Function that reads all filtered cells from a list object and returns a two-dimensional variant array with filter data
    'Arr structure will be as follows (dimension limits displayed near to the fields)
    '
    '(0,0)  Filter1:    field   criteria1   op  criteria2   (0,3)
    '(1,0)  Filter2:    field   criteria1   op  criteria2   (1,3)
    ' ...
    '
    'Every filter field will be checked if it is active. Only active filters' data will be copied.
    '
    'Input args:
    '   lo:             The list object containing the filters (if any)
    '
    'Output args:
    '   GetAutoFilters: The multi-dimensional filter data array
    
    'Init output
    Dim filterData() As Variant
    GetAutoFilters = filterData
    
    Dim filt As Filters: Set filt = lo.AutoFilter.Filters
    Dim fCount As Long: fCount = filt.Count

    
    Dim fIdx As Long
    Dim cFilter As Filter
    Dim onFilterFields() As Long
    
    For fIdx = 0 To filt.Count - 1
        'Cycle through all filter fields and store the field numbers if filter is set to 'on'
        Set cFilter = filt(fIdx + 1)
        If cFilter.On Then
            If Not Base.IsArrayAllocated(onFilterFields) Then
                'Init the array containing the field numbers
                ReDim onFilterFields(0)
            Else
                'Increase the size of the array every cycle
                ReDim Preserve onFilterFields(UBound(onFilterFields) + 1)
            End If
            'Set the field number (filter items are counted from 1 to n), adjust the index
            onFilterFields(UBound(onFilterFields)) = fIdx + 1
        End If
    Next fIdx
    
    If Not Base.IsArrayAllocated(onFilterFields) Then Exit Function
    
    'Init the filter data array
    ReDim filterData(UBound(onFilterFields), 0 To 3)
    Dim fieldIdx As Long
    Dim cField As Long
    For fieldIdx = 0 To UBound(onFilterFields)
        'Cycle through all active filters and get their data
        cField = onFilterFields(fieldIdx)
        Set cFilter = filt(cField)

        Dim critTwoIsSet As Boolean: critTwoIsSet = False
        Dim opIsSet As Boolean: opIsSet = False
        Dim tester As Variant
                
        'Use error handler to detect whether the operator or the criteria2 property for a filter is set or not (error if an unset property is read)
        On Error Resume Next
        tester = cFilter.Operator
        If Err.Number = 0 Then
            'No error reading the property
            opIsSet = True
        End If
        Err.Clear
        tester = cFilter.Criteria2
        If Err.Number = 0 Then
            'No error reading the property
            critTwoIsSet = True
        End If
        Err.Clear
        On Error GoTo 0 'Reset error handling
            
        filterData(fieldIdx, 0) = cField
        filterData(fieldIdx, 1) = cFilter.Criteria1
        
        'Only store operator and criteria2 field if they are set
        If opIsSet Then filterData(fieldIdx, 2) = cFilter.Operator
        If critTwoIsSet Then filterData(fieldIdx, 3) = cFilter.Criteria2
    Next fieldIdx
    
    GetAutoFilters = filterData
End Function



Function ReapplyAutoFilters(ByRef lo As ListObject, filterData() As Variant)
    'Reapply the stored data from a multi-dimensional filter data array.
    'For further details see 'GetAutoFilter' function
    '
    'Input args:
    '   lo:         The list object to apply the filters to
    '   filterData: The filter data that was stored and is now set again
    
    Dim fIdx As Long

    For fIdx = 0 To UBound(filterData)
        'For every stored filter cycle
        
        'Check if operator and / or criteria2 is set
        Dim opIsSet As Boolean: opIsSet = (filterData(fIdx, 2) <> 0)
        Dim critTwoIsSet As Boolean: critTwoIsSet = Not IsEmpty(filterData(fIdx, 3))
        
        'Apply the filters depending on which data is available
        If critTwoIsSet And opIsSet Then
            'Criteria2 is set (that means op is set as well)
            Call lo.Range.AutoFilter(filterData(fIdx, 0), filterData(fIdx, 1), filterData(fIdx, 2), filterData(fIdx, 3))
        ElseIf opIsSet Then
            Call lo.Range.AutoFilter(filterData(fIdx, 0), filterData(fIdx, 1), filterData(fIdx, 2))
        Else
            Call lo.Range.AutoFilter(filterData(fIdx, 0), filterData(fIdx, 1))
        End If
    Next fIdx
End Function



Function DeleteWorksheetSilently(sheet As Worksheet)
    'Delete a worksheet without warning. Make sure to unhide it first. Otherwise deletion will fail
    '
    'Input args:
    '   sheet:  Sheet you want to delete
    
    'Check args
    If sheet Is Nothing Then Exit Function
    
    'Delete the sheet
    Application.DisplayAlerts = False
    sheet.Visible = xlSheetVisible 'Make visible prior to deletion to prevent errors
    sheet.Delete
    Application.DisplayAlerts = True
End Function

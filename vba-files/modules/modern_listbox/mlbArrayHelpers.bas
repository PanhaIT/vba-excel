Attribute VB_Name = "mlbArrayHelpers"
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM

' ARRAY HELPER FUNCTIONS
Public Function IsArrayAllocated(ByRef arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And (Not IsError(LBound(arr, 1))) And LBound(arr, 1) <= UBound(arr, 1)
End Function

Public Function arrayGetRowCount(data As Variant) As Long
    If IsEmpty(data) = False Then
        arrayGetRowCount = UBound(data, 1) - LBound(data, 1) + 1
    End If
End Function

Public Function GetRow(ByRef data As Variant, ByVal row As Long) As Variant
    Dim temp As Variant
    Call arraySetSize(temp, data, 1)
    Call arrayCopyRow(temp, 1, data, row)
    GetRow = temp
End Function

Public Function arrayGetColumnCount(ByRef data As Variant) As Long
    arrayGetColumnCount = 0
    If IsEmpty(data) = False Then
        arrayGetColumnCount = UBound(data, 2) - LBound(data, 2) + 1
    End If
End Function

' Name: arraySetSize
' Description: Sets the size of the destinationArray to the size of the sourceArray
Public Sub arraySetSize(ByRef destinationArray As Variant _
                        , ByRef sourceArray As Variant _
                        , Optional ByVal rowCount As Long = -1)
    If rowCount = -1 Then
        rowCount = UBound(sourceArray, 1)
     End If
    ReDim destinationArray(LBound(sourceArray, 1) To rowCount _
                            , LBound(sourceArray, 2) To UBound(sourceArray, 2))

End Sub

Public Function arrayGetSize(ByRef arr() As Long) As Long
    arrayGetSize = UBound(arr) - LBound(arr) + 1
End Function

Public Sub arrayCopy(ByRef destinationArray As Variant _
                        , ByRef sourceArray As Variant _
                        , Optional ByVal rowCount As Long = -1)
    If rowCount = -1 Then
        rowCount = UBound(sourceArray, 1)
    End If
     
    Dim i As Long, j As Long
    For i = LBound(sourceArray, 1) To rowCount
        For j = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            destinationArray(i, j) = sourceArray(i, j)
            debug.print "arrayCopy=> row: " & i & ", column: " & j & " = " & sourceArray(i, j)
        Next j
    Next i
End Sub

Public Sub arrayCopyRow(ByRef destinationArray As Variant _
                        , ByRef destinationRow As Long _
                        , ByRef sourceArray As Variant _
                        , ByRef sourceRow As Long _
                        , Optional ByRef startColumn As Long = 1)
     
    Dim Start As Long
    Start = LBound(sourceArray, 2) + (startColumn - 1)
    
    Dim column As Long: column = 1
    Dim j As Long
    For j = Start To UBound(sourceArray, 2)
        'copy a selected or all selected rows from listbox selected row to the excel sheet
        destinationArray(destinationRow, column) = sourceArray(sourceRow, j)
          ' debug.Print "arrayCopyRow=>sourceRow: " & sourceRow & ", column: " & j & " = " & sourceArray(sourceRow, j)
        column = column + 1
    Next j
    
End Sub

Public Function arrayToString(arr As Variant) As String
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        arrayToString = arrayToString & arrayRowToString(arr, i) & vbNewLine
        ' debug.Print "row: " & i & ",  arrayToString=> = " & arrayRowToString(arr, i)
    Next i
End Function

Public Function arrayRowToString(arr As Variant, row As Long) As String
    Dim i As Long
    For i = LBound(arr, 2) To UBound(arr, 2)
        arrayRowToString = arrayRowToString & arr(row, i) & " "
        debug.Print "arrayRowToString=> row:" & row & ", Column: " & i & " = " & arr(row, i)
    Next i
End Function

Public Sub ArrayToRange(ByRef arr As Variant, rg As Range)
    rg.Resize(UBound(arr, 1) - LBound(arr, 1) + 1, UBound(arr, 2) - LBound(arr, 2) + 1) = arr
End Sub

Public Sub ArrayToRangePartial(ByRef arr As Variant, rg As Range, ByVal rowCount As Long)
    rg.Resize(rowCount, UBound(arr, 2) - LBound(arr, 2) + 1) = arr
End Sub

Public Sub ArrayToRangeFormula(ByRef arr As Variant, rg As Range)
    rg.Resize(UBound(arr, 1) - LBound(arr, 1) + 1, UBound(arr, 2) - LBound(arr, 2) + 1).Formula2 = arr
End Sub

Public Sub arrayCopyNonSequentialRows(ByRef arr As Variant, ByRef rowNumbers As Object, ByRef rowsOut As Variant)
    Debug.Assert IsEmpty(rows()) = False And IsEmpty(arr) = False
    
    ReDim rowsOut(1 To rowNumbers.Count, 1 To arrayGetColumnCount(arr))
    Dim rowCount As Long: rowCount = 1
    Dim key As Variant
    For Each key In rowNumbers.Keys
        Call arrayCopyRow(rowsOut, rowCount, arr, CLng(key))
        rowCount = rowCount + 1
    Next key
End Sub
    
Public Function StringTo2DArray(text As String, delimeter As String) As Variant
    Dim arr As Variant: arr = Split(text, delimeter)
    ReDim arrNew(1 To 1, 1 To UBound(arr) - LBound(arr) + 1) As Variant
    
    Dim i As Long, column As Long: column = 1
    For i = LBound(arr) To UBound(arr)
        arrNew(1, column) = arr(i)
        column = column + 1
    Next i
    StringTo2DArray = arrNew
End Function

Public Function SortDictionaryByValue(d As Object, Optional sortOrder As XlSortOrder = xlAscending) As Dictionary
        
    ' Set the return dictionary
    Set SortDictionaryByValue = d
            
    ' Not need to sort if less that two items
    If d.Count <= 1 Then Exit Function
    
    ' Convert to sort function order type
    Dim order As Long: order = IIf(sortOrder = xlAscending, 1, -1)

    ' Create array to store values so we can sort them
    Dim data As Variant
    Debug.Assert d.Count > 0
    ReDim data(1 To d.Count, 1 To 2) As Variant
    
    Dim row As Long: row = 1
    Dim key As Variant
    For Each key In d.Keys
        data(row, 1) = key
        data(row, 2) = d(key)
        row = row + 1
    Next
    
    ' Sort by value
    data = WorksheetFunction.Sort(data, 2, order)
    
    ' Remove all exist values
    d.RemoveAll
    
    ' Repopulate dictionary from sorted array
    Dim i As Long
    For i = LBound(data) To UBound(data)
        d.Add data(i, 1), data(i, 2)
    Next i
    

 End Function





















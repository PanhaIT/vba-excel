Attribute VB_Name = "mdlSorting"

' Public arrIndex As Integer

' Function SelectionSort(arr As Variant, _
'                        Optional colIndex As Integer = 0)             
'     Dim i As Long
'     Dim j As Long
'     Dim sTemp As Variant
'     Dim k  As Integer
    
'     Select Case SortDirection
'         Case Is = "asc"
'             For i = LBound(arr) To UBound(arr)
'                 For j = i + 1 To UBound(arr)
'                     If arr(i, colIndex) > arr(j, colIndex) Then
'                         For k = 0 To colCount
'                             sTemp = arr(i, k)
'                             arr(i, k) = arr(j, k)
'                             arr(j, k) = sTemp
'                         Next k
        
'                     End If
'                 Next j
'             Next i
        
'         Case Is = "desc"
'             For i = LBound(arr) To UBound(arr)
'                 For j = i + 1 To UBound(arr)
'                     If arr(i, colIndex) < arr(j, colIndex) Then
'                         For k = 0 To colCount
'                             sTemp = arr(i, k)
'                             arr(i, k) = arr(j, k)
'                             arr(j, k) = sTemp
'                             Debug.Print "desc=>" & arr(j, k)
'                         Next k
        
'                     End If
'                 Next j
'             Next i
  
'     End Select
'     SelectionSort = arr
' End Function




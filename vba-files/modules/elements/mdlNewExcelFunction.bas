Attribute VB_Name ="mdlNewExcelFunction"

Public Function Plookup(LookupVal As Integer, TableArray As Range, ColumnIndex As Integer,Optional CriteriaRange1 As Range, Optional Criteria1 As Integer) As Double
    Dim item_row As Range
    Dim i As Integer
    Dim PlookupTmp As Double
  
    PlookupTmp = 0
    For i = 1 To TableArray.Rows.Count
        If TableArray.Cells(i, 1) <> "" Then
            If (TableArray.Cells(i, 1) = LookupVal AND CriteriaRange1.Cells(i, 1)=Criteria1) Then
                PlookupTmp = TableArray.Cells(i, ColumnIndex)
            End If
        End If
    Next i
    Plookup = PlookupTmp
End Function

Public Function Slookup(LookupVal As Integer, TableArray As Range, ColumnIndex As Integer, Optional CriteriaRange1 As Range, Optional Criteria1 As Integer, Optional CriteriaRange2 As Range, Optional Criteria2 As Integer) As Double
    Dim i As Integer
    Dim SlookupTmp As Double
    
    SlookupTmp = 0
    For i = 1 To TableArray.Rows.Count
        If TableArray.Cells(i, 1) <> "" Then
            If (TableArray.Cells(i, 1) = LookupVal And CriteriaRange1.Cells(i, 1) = Criteria1 And CriteriaRange2.Cells(i, 1) = Criteria2) Then
                SlookupTmp = TableArray.Cells(i, ColumnIndex)
            End If
        End If
    Next i
    Slookup = SlookupTmp
End Function

Public Function Zlookup(LookupVal As Integer, TableArray As Range, ColumnIndex As Integer, Optional CriteriaRange1 As Range, Optional Criteria1 As Integer, Optional CriteriaRange2 As Range, Optional Criteria2 As Integer, Optional CriteriaRange3 As Range, Optional Criteria3 As Integer) As Double
    Dim i As Integer
    Dim ZlookupTmp As Double

    ZlookupTmp = 0
    For i = 1 To TableArray.Rows.Count
        If TableArray.Cells(i, 1) <> "" Then
            If (TableArray.Cells(i, 1) = LookupVal And CriteriaRange1.Cells(i, 1) = Criteria1 And CriteriaRange2.Cells(i, 1) = Criteria2 And CriteriaRange3.Cells(i, 1) = Criteria3) Then
                ZlookupTmp = TableArray.Cells(i, ColumnIndex)
            End If
        End If
    Next i
    Zlookup = ZlookupTmp
End Function
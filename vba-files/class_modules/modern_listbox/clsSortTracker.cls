VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSortTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Tracks which column is currently sorted and it's sorted status(asc or desc)

Private m_ColumnSortStatus() As Long

Public Function SetSortDirection(ByVal column As Long) As XlSortOrder

    If m_ColumnSortStatus(column) = 0 Then
        m_ColumnSortStatus(column) = 1
    Else
        m_ColumnSortStatus(column) = IIf(m_ColumnSortStatus(column) = 1, -1, 1)
    End If
    SetSortDirection = m_ColumnSortStatus(column)
    
End Function

Public Function GetSortDirection(ByVal column As Long) As XlSortOrder
    GetSortDirection = m_ColumnSortStatus(column)
End Function

Public Sub SetColumnSortStatus(ColumnCount As Long)

    ReDim m_ColumnSortStatus(1 To ColumnCount)
    Dim i As Long
    For i = LBound(m_ColumnSortStatus) To UBound(m_ColumnSortStatus)
        m_ColumnSortStatus(i) = -1
    Next i
End Sub

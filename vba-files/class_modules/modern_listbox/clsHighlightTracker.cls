VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHighlightTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Tracks the rows which are currently selected


Private Const NO_ROW_SET As Long = -1
Private m_SelectedItems As Object
' Private m_SelectedItemsDictionary As Dictionary
' Private m_SelectedItemsDictionaryKeys As Object
' Private m_SelectedItemsDictionaryKeysCount As Long

Public Property Get SelectedRowNumbers() As Object
    Set SelectedRowNumbers = CreateRowNumberDictionary
    Debug.Assert Not SelectedRowNumbers Is Nothing
End Property

' Returns the first selected row
Public Property Get SelectedRow() As Long
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    SelectedRow = m_SelectedItems.Keys(0)
    Debug.Assert SelectedRow > 0
End Property

Public Property Get SelectedRowCount() As Long
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    SelectedRowCount = m_SelectedItems.Count
End Property

' Returns the first selected row
Public Property Get SelectedFirstVisibleRow() As Long
    Set m_SelectedItems = SortDictionaryByValue(m_SelectedItems, xlAscending)
    Dim key As Variant
    For Each key In m_SelectedItems.Keys
        If m_SelectedItems(key) <> NO_ROW_SET Then
            SelectedFirstVisibleRow = m_SelectedItems(key)
            Exit For
        End If
    Next key
End Property

Public Property Get SelectedVisibleRowCount() As Long
    Dim key As Variant
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    For Each key In m_SelectedItems.Keys
        If m_SelectedItems(key) <> NO_ROW_SET Then
            SelectedVisibleRowCount = SelectedVisibleRowCount + 1
        End If
    Next key
End Property

Public Sub ClearAll()
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    m_SelectedItems.RemoveAll
End Sub
Public Sub SelectAll(data As Variant, recordCount As Long)
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    m_SelectedItems.RemoveAll
    Dim i As Long, rowID As Long
    For i = 1 To recordCount
        rowID = data(i, 1)
        m_SelectedItems.Add rowID, i
    Next i
End Sub

Public Sub SortSelectedItems(sortOrder As XlSortOrder)
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    If m_SelectedItems.Count > 1 Then
        Set m_SelectedItems = SortDictionaryByValue(m_SelectedItems, sortOrder)
    End If
End Sub

Public Function IsRowHighlighted(rowID As Long) As Boolean
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    IsRowHighlighted = m_SelectedItems.Exists(rowID)
End Function

Public Sub SelectRow(record As recordDetails, multiSelect As fmMultiSelect)
    
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    If multiSelect = fmMultiSelectSingle Then
        Call selectRowSingle(record)
    Else
        Call selectRowMulti(record)
    End If
End Sub

' Called to select/deselect a row . This is when the method SetSelected is called by the UserForm code
Public Sub SetRow(record As recordDetails, multiSelect As fmMultiSelect, turnOn As Boolean)
    If multiSelect = fmMultiSelectSingle Then
        Call setRowSingle(record, turnOn)
    Else
        Call setRowMulti(record, turnOn)
    End If
End Sub

' Used after sort or filter operations to update the new row positions for the records
Public Sub ResetRecordIDs(data As Variant)
    ' Dim existingIDs As New Dictionary
    Dim existingIDs As Object
    Set existingIDs = CreateObject("Scripting.Dictionary")
    Set m_selecteditems = CreateObject("Scripting.Dictionary")

    Dim row As Long, rowID As Long
    For row = 1 To UBound(data, 1)
        rowID = data(row, 1)
        ' If ID is already selected then update with new row position
        If m_SelectedItems.Exists(rowID) Then
            existingIDs.Add rowID, 0
            m_SelectedItems(rowID) = row
        End If
    Next row
    
    ' Records that no longer exist in the current data have their row
    ' reset to avoid that row being selected
    Dim ID As Variant
    For Each ID In m_SelectedItems.Keys
        If existingIDs.Exists(ID) = False Then
            m_SelectedItems(ID) = NO_ROW_SET
        End If
    Next ID
End Sub

' Private subs
Private Function CreateRowNumberDictionary() As Object
    Dim rowNumbers As Object
    
    Set rowNumbers      = CreateObject("Scripting.Dictionary")
    Set m_selecteditems = CreateObject("Scripting.Dictionary")

    Dim key As Variant
    For Each key In m_SelectedItems.Keys
        If m_SelectedItems(key) <> NO_ROW_SET Then
            rowNumbers.Add m_SelectedItems(key), key
        End If
    Next key
    Set CreateRowNumberDictionary = rowNumbers
End Function

Private Sub selectRowSingle(record As recordDetails)
    ' MsgBox "record.ID=" & record.ID,,"selectRowSingle" 

    If m_SelectedItems.Exists(record.ID) = True Then
        Call m_SelectedItems.Remove(record.ID)
    Else
        Call ClearAll
        Call m_SelectedItems.Add(record.ID, record.realRowPosition)
    End If
End Sub

Private Sub selectRowMulti(record As recordDetails)
    If m_SelectedItems.Exists(record.ID) = True Then
        Call m_SelectedItems.Remove(record.ID)
    Else
        MsgBox "record.ID=" & record.ID & " , check exist=" &  m_SelectedItems.Exists(record.ID) & ", record.realRowPosition=" & record.realRowPosition
        Call m_SelectedItems.Add(record.ID, record.realRowPosition)
    End If
End Sub

Private Sub setRowSingle(record As recordDetails, turnOn As Boolean)
    Set m_selecteditems = CreateObject("Scripting.Dictionary")
    If turnOn = True And m_SelectedItems.Exists(record.ID) = False Then
        ' Only turn on if not selected already
        Call ClearAll
        Call m_SelectedItems.Add(record.ID, record.realRowPosition)
    End If
    If turnOn = False And m_SelectedItems.Exists(record.ID) = True Then
        ' Only turn off if selected already
        Call m_SelectedItems.Remove(record.ID)
    End If
End Sub

Private Sub setRowMulti(record As recordDetails, turnOn As Boolean)
    If turnOn = True And m_SelectedItems.Exists(record.ID) = False Then
        ' Only turn on if not selected already
        Call m_SelectedItems.Add(record.ID, record.realRowPosition)
    End If
    If turnOn = False And m_SelectedItems.Exists(record.ID) = True Then
        ' Only turn off if selected already
        Call m_SelectedItems.Remove(record.ID)
    End If
End Sub




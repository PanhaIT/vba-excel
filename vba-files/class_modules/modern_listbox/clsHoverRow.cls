VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHoverRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Manages the hover highlight of the row

'Private m_parentFrame As MSForms.frame
Private m_parentFrame As clsParentFrameManager
Private m_settings As clsListBoxSettings
Private m_controlAccessors As New clsControlAccessors
Private m_controlNameHelpers As New clsControlNameHelpers
Private m_currentHoverRow As Long
Private m_hightlightedHeaderColumn As Long
Private m_InitializedCorrectly As Boolean

Public Property Let InitializedCorrectly(newInitializedCorrectly As Boolean)
    m_InitializedCorrectly = True
End Property
Public Property Set parentFrame(newParentFrame As clsParentFrameManager)
    Set m_parentFrame = newParentFrame
    Set m_controlAccessors.parentFrame = newParentFrame
End Property
Public Property Set settings(argSettings As clsListBoxSettings)
    Set m_settings = argSettings
End Property
Private Sub Class_Initialize()
    Call ClearRowTracker
End Sub

Public Sub ClearHover()
    Call ClearRowHoverControls
    Call ClearRowTracker
End Sub

Public Sub SetRowHover(ByVal row As Long)
    If IsCurrentRowHover(row) = True Then Exit Sub
    Call ClearRowHoverControls
    Call SetRowTracker(row)
    Call SetRowHoverControls
End Sub

' Check that row has hover hightlight. It may have selection hightlight and we don't want to clear that
Private Function IsRowColorHighlight() As Boolean
    Dim frame As MSForms.frame
    Set frame = m_parentFrame.parentFrame.Controls(m_controlNameHelpers.BuildRowFrameName(m_currentHoverRow))
    IsRowColorHighlight = frame.backcolor = m_settings.hoverColorRow()
End Function

Private Sub SetRowHoverControls()

    ' Set the frame back color
    Dim frame As MSForms.frame
    Set frame = m_controlAccessors.GetRowFrame(m_currentHoverRow)
    frame.backcolor = m_settings.hoverColorRow()

    ' Set label backcolor
    Dim currentLabel As MSForms.Control
    For Each currentLabel In frame.Controls
        If m_controlNameHelpers.GetControlTypename(currentLabel.name) = "label" Then
            currentLabel.backcolor = m_settings.hoverColorRow
        End If
    Next

End Sub



Private Sub ClearRowHoverControls()
    Debug.Assert Not m_parentFrame Is Nothing
    ' Clear the back color of highlighted Row frame
    
    If IsAnyRowHighlighted() = False Then Exit Sub
    
    Dim frame As MSForms.frame
    Set frame = m_controlAccessors.GetRowFrame(m_currentHoverRow)
    frame.backcolor = m_settings.getRowBackgroundColor()

    Dim currentLabel As MSForms.Control
    For Each currentLabel In frame.Controls
        
        If m_controlNameHelpers.GetControlTypename(currentLabel.name) <> "label" Then GoTo continue

        ' Clear label
        currentLabel.backcolor = m_settings.getRowBackgroundColor()

continue:
    Next currentLabel
End Sub


' Highlighting - Helper procedures

' If row is selected then we don't highlight it with hover color
'Private Function IsRowSelected(ByVal row As Long, SelectedItems As Collection) As Boolean
'
'    If SelectedItems.Count = 0 Then Exit Function
'    Dim currentRow As Variant
'    For Each currentRow In SelectedItems
'        If currentRow = row Then
'            IsRowSelected = True
'            Exit Function
'        End If
'    Next
'End Function

Private Function IsCurrentRowHover(row As Long) As Boolean
   IsCurrentRowHover = (row = m_currentHoverRow)
End Function

Private Function IsAnyRowHighlighted() As Boolean
   IsAnyRowHighlighted = (m_currentHoverRow <> -1)
End Function
Private Sub ClearRowTracker()
    m_currentHoverRow = -1
End Sub
Private Sub SetRowTracker(ByVal row As Long)
    m_currentHoverRow = row
End Sub


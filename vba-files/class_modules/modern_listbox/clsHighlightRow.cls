VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHighlightRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Highlights the row

'Private m_parentFrame As MSForms.frame
Private m_parentFrame As clsParentFrameManager
Private m_settings As clsListBoxSettings
Private m_controlAccessors  As New clsControlAccessors
Private m_controlNameHelpers As New clsControlNameHelpers
Private m_InitializedCorrectly As Boolean

Private Sub Class_Initialize()
    m_InitializedCorrectly = False
End Sub

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

Public Sub ApplyHighlight(visibleRecords As visibleRecordRange, rowNumbers As Object) 'Dictionary
    Debug.Assert Not rowNumbers Is Nothing
    Set rowNumbers  = CreateObject("Scripting.Dictionary")

    Dim row As Long: row = 1
    Dim i As Long
    For i = visibleRecords.Start To visibleRecords.end
        debug.print "rowNumbers.Exists(i) =" & rowNumbers.Exists(i) 
        If rowNumbers.Exists(i) Then
            Call SetCheck(row, checkOn:=True)
            Call SetRowHighlight(row, turnOn:=True)
        Else
            Call SetCheck(row, checkOn:=False)
            Call SetRowHighlight(row, turnOn:=False)
        End If
        row = row + 1
    Next i
End Sub
'


Public Sub SetCheckHeader(ByVal checkOn As Boolean)
    m_controlAccessors.GetCheckLabelHeader().caption = IIf(checkOn = True, ChrW(80), "") ' myicon checkbox => ChrW(&HE041)
End Sub

Private Sub SetCheck(ByVal row As Long, ByVal checkOn As Boolean)
    Debug.Assert row > 0
    m_controlAccessors.GetCheckLabel(row).caption = IIf(checkOn = True, ChrW(80), "") ' myicon checkbox => ChrW(&HE041)
End Sub

Private Sub SetRowHighlight(ByVal row As Long, turnOn As Boolean)
    
    ' Set the frame back color
    Dim frame As MSForms.frame
    Set frame = m_parentFrame.parentFrame().Controls(m_controlNameHelpers.BuildRowFrameName(row))
    frame.backcolor = _
                    IIf(turnOn = True, m_settings.rowSelectionColor(), m_settings.getRowBackgroundColor())
    
    ' Set the labels back color
    Dim currentLabel As MSForms.Label
   
    Dim currentControl As MSForms.Control
    For Each currentControl In frame.Controls
        If m_controlNameHelpers.IsRowLabel(currentControl.name) Then
            currentControl.backcolor = _
                    IIf(turnOn = True, m_settings.rowSelectionColor(), m_settings.getRowBackgroundColor())
        End If
    Next

End Sub















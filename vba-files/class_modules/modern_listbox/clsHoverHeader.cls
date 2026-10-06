VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHoverHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Manages the header hover highlight

'Private m_parentFrame As MSForms.frame
'Private m_parentFrame As clsParentFrameManager
Private m_settings As clsListBoxSettings
Private m_currentHoverColumn As Long
Private m_controlAccessors As New clsControlAccessors
Private m_InitializedCorrectly As Boolean

Public Property Let InitializedCorrectly(newInitializedCorrectly As Boolean)
    m_InitializedCorrectly = True
End Property
Public Property Set parentFrame(newParentFrame As clsParentFrameManager)
  '  Set m_parentFrame = newParentFrame
    Set m_controlAccessors.parentFrame = newParentFrame
End Property
Public Property Set settings(argSettings As clsListBoxSettings)
    Set m_settings = argSettings
End Property

Private Sub Class_Initialize()
    Call ClearColumnTracker
     m_InitializedCorrectly = False
End Sub

Public Sub SetHover(ByVal column As Long)

    If IsColumnAlreadyHover(column) = True Then Exit Sub
    
    Call ClearHover
    Call SetColumnHover(column)
    Call SetColumnTracker(column)

End Sub

Public Sub ClearHover()
    Call ClearColumnHover
    Call ClearColumnTracker
End Sub

Private Sub ClearColumnHover()
    If IsAnyColumnHover() = False Then Exit Sub
    m_controlAccessors.GetHeaderLabel(m_currentHoverColumn).backcolor = m_settings.columnHeaderBackgroundColor
    m_controlAccessors.GetHeaderLabelFrame(m_currentHoverColumn).backcolor = m_settings.columnHeaderBackgroundColor
    m_controlAccessors.GetHeaderLabelSymbol(m_currentHoverColumn).backcolor = m_settings.columnHeaderBackgroundColor
End Sub

Private Sub SetColumnHover(column As Long)
    
    m_controlAccessors.GetHeaderLabel(column).backcolor = m_settings.hoverColorheader
    m_controlAccessors.GetHeaderLabelFrame(column).backcolor = m_settings.hoverColorheader
    m_controlAccessors.GetHeaderLabelSymbol(column).backcolor = m_settings.hoverColorheader

End Sub

Private Function IsColumnAlreadyHover(ByVal column As Long) As Boolean
    IsColumnAlreadyHover = (m_currentHoverColumn = column)
End Function

Private Function IsAnyColumnHover() As Boolean
    IsAnyColumnHover = (m_currentHoverColumn <> -1)
End Function

Private Sub SetColumnTracker(ByVal column As Long)
    m_currentHoverColumn = column
End Sub

Private Sub ClearColumnTracker()
    m_currentHoverColumn = -1
End Sub

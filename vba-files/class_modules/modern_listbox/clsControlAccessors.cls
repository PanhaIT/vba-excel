VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlAccessors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Helper procedures for accessing the different controls
Private m_controlNameHelpers As New clsControlNameHelpers
Private m_parentFrame As clsParentFrameManager

Public Property Set parentFrame(newParentFrame As clsParentFrameManager)
    Set m_parentFrame = newParentFrame
End Property

Public Function GetCheckLabelHeader() As MSForms.Label
    Set GetCheckLabelHeader = m_parentFrame.parentFrame.Controls(m_controlNameHelpers.BuildCheckboxLabelNameHeader)
    Debug.Assert Not GetCheckLabelHeader Is Nothing
End Function

Public Function GetCheckFrameHeader() As MSForms.frame
    Set GetCheckFrameHeader = m_parentFrame.parentFrame.Controls(m_controlNameHelpers.BuildCheckboxNameHeader())
    Debug.Assert Not GetCheckFrameHeader Is Nothing
End Function

Public Function GetCheckLabel(row As Long) As MSForms.Label
    Set GetCheckLabel = GetCheckFrame(row).Controls(m_controlNameHelpers.BuildCheckboxLabelName(row))
    Debug.Assert Not GetCheckLabel Is Nothing
End Function

Public Function GetCheckFrame(row As Long) As MSForms.frame
    Set GetCheckFrame = GetRowFrame(row).Controls(m_controlNameHelpers.BuildCheckboxName(row))
    Debug.Assert Not GetCheckFrame Is Nothing
End Function

Public Function GetRowFrame(row As Long) As MSForms.frame
    Set GetRowFrame = m_parentFrame.parentFrame.Controls(m_controlNameHelpers.BuildRowFrameName(row))
    Debug.Assert Not GetRowFrame Is Nothing
End Function

Public Function GetHeaderFrame() As MSForms.frame
    Set GetHeaderFrame = m_parentFrame.parentFrame.Controls(m_controlNameHelpers.BuildHeaderFrameName())
    Debug.Assert Not GetHeaderFrame Is Nothing
End Function

Public Function GetRowLabel(row As Long, column As Long) As MSForms.Label
    Set GetRowLabel = GetRowFrame(row).Controls(m_controlNameHelpers.BuildRowLabelName(row, column))
    Debug.Assert Not GetRowLabel Is Nothing
End Function

Public Function GetHeaderLabel(column As Long) As MSForms.Label
    Set GetHeaderLabel = m_parentFrame.parentFrame.Controls(m_controlNameHelpers.BuildHeaderLabelName(column))
End Function

Public Function GetHeaderLabelFrame(column As Long) As MSForms.frame
    Set GetHeaderLabelFrame = m_parentFrame.parentFrame.Controls(m_controlNameHelpers.BuildHeaderLabelFrameName(column))
End Function

Public Function GetHeaderLabelSymbol(column As Long) As MSForms.Label
    Set GetHeaderLabelSymbol = m_parentFrame.parentFrame.Controls(m_controlNameHelpers.BuildLabelSymbolName(column))
End Function
      
Public Function GetPaginationNumberLabel(ID As Long, column As Long) As MSForms.Label
    Set GetPaginationNumberLabel = m_parentFrame.parentFrame.Parent.Controls(m_controlNameHelpers.BuildPaginationNumberLabelName(ID, column))
End Function

Public Function GetSearchBox(ID As Long) As MSForms.TextBox
    Set GetSearchBox = m_parentFrame.parentFrame.Parent.Controls(m_controlNameHelpers.BuildSearchTextboxName(ID))
End Function

Public Function CheckSearchBoxExists(ID As Long) As Boolean
    On Error Resume Next
    Dim searchBox As MSForms.TextBox
    Set searchBox = m_parentFrame.parentFrame.Parent.Controls(m_controlNameHelpers.BuildSearchTextboxName(ID))
    CheckSearchBoxExists = Not searchBox Is Nothing
End Function

Public Function CheckPaginationExists(ID As Long) As Boolean
    On Error Resume Next
    Dim searchBox As MSForms.TextBox
    Set searchBox = m_parentFrame.parentFrame.Parent.Controls(m_controlNameHelpers.BuildPaginationArrowLabelName(ID, 1))
    CheckPaginationExists = Not searchBox Is Nothing
End Function

Public Function GetSearchLabel(ID As Long) As MSForms.Label
    Set GetSearchLabel = m_parentFrame.parentFrame.Parent.Controls(m_controlNameHelpers.BuildSearchLabelName(ID))
End Function
Public Function GetPaginationArrowLabel(ID As Long, arrowNumber As Long) As MSForms.Label
    Set GetPaginationArrowLabel = m_parentFrame.parentFrame.Parent.Controls(m_controlNameHelpers.BuildPaginationArrowLabelName(ID, arrowNumber))
End Function



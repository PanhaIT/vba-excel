VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlNameHelpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Helper procedures to create the names for the controls that are part of the Listbox

Public Function BuildCheckboxLabelName(ByVal row As Long) As String
   BuildCheckboxLabelName = "checkboxLabel" & "-" & row
End Function

Public Function BuildCheckboxLabelNameHeader() As String
   BuildCheckboxLabelNameHeader = "checkboxLabelHeader"
End Function

Public Function BuildCheckboxName(ByVal row As Long) As String
   BuildCheckboxName = "checkbox" & "-" & row
End Function

Public Function BuildCheckboxNameHeader() As String
   BuildCheckboxNameHeader = "checkboxHeader"
End Function

Public Function BuildHeaderFrameName() As String
   BuildHeaderFrameName = "HeaderRowFrame"
End Function

Public Function BuildHeaderLabelFrameName(ByVal column As Long) As String
   BuildHeaderLabelFrameName = "HeaderFrameLabel" & "-" & column
End Function

Public Function BuildHeaderLabelName(ByVal column As Long) As String
   BuildHeaderLabelName = "headerLabel" & "-" & column
End Function

Public Function BuildLabelSymbolName(ByVal column As Long) As String
   BuildLabelSymbolName = "labelSymbol" & "-" & column
End Function

Public Function BuildPaginationArrowLabelName(ID As Long, ByVal number As Long) As String
   BuildPaginationArrowLabelName = "pageArrow" & ID & "-" & number
End Function

' 07-05-2025 Added ID
Public Function BuildPaginationNumberLabelName(ID As Long, ByVal number As Long) As String
   BuildPaginationNumberLabelName = "pageNumber" & ID & "-" & number
End Function

Public Function BuildRowFrameName(ByVal row As Long) As String
   BuildRowFrameName = "rowFrame" & "-" & row
End Function

Public Function BuildRowLabelName(ByVal row As Long, ByVal column As Long) As String
   BuildRowLabelName = "label" & "-" & row & "-" & column
End Function

Public Function BuildSearchLabelName(ID As Long) As String
   BuildSearchLabelName = "labelSearch" & CStr(ID)
End Function

Public Function BuildSearchTextboxName(ID As Long) As String
   BuildSearchTextboxName = "textboxSearch" & CStr(ID)
End Function

Public Function GetColumnFromHeaderFrameName(name As String) As Long
   GetColumnFromHeaderFrameName = Split(name, "-")(1)
End Function

Public Function GetColumnFromHeaderLabelName(name As String) As String
   GetColumnFromHeaderLabelName = Split(name, "-")(1)
End Function

Public Function GetColumnFromLabelSymbolName(name As String) As String
   GetColumnFromLabelSymbolName = CLng(Split(name, "-")(1))
End Function

Public Function GetColumnFromRowLabelName(name As String) As String
   GetColumnFromRowLabelName = Split(name, "-")(2)
End Function

Public Function GetControlTypename(name As String) As String
   GetControlTypename = Split(name, "-")(0)
End Function

Public Function GetLabelEventNumber(row As Long, column As Long, ColumnCount As Long) As Long
    GetLabelEventNumber = (row * ColumnCount) + column
End Function

Public Function GetRowFromFrameName(name As String) As Long
   GetRowFromFrameName = Split(name, "-")(1)
End Function

Public Function IsRowLabel(ByVal name As String) As Boolean
   IsRowLabel = Split(name, "-")(0) = "label"
End Function


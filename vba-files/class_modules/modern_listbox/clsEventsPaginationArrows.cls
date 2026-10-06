VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsPaginationArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Events for the pagination arrows

Private m_modernListbox As clsModernListbox
Private m_brandModernListbox As clsModernListbox
Public WithEvents mylabel As MSForms.Label
Public WithEvents myBrandLabel As MSForms.Label

Public Property Set parentListBox(ByVal argModernListBox As clsModernListbox)
    Set m_modernListbox = argModernListBox
End Property

Public Property Set parentBrandListBox(ByVal argBrandModernListBox As clsModernListbox)
    Set m_brandModernListbox = argBrandModernListBox
End Property

Private Sub mylabel_Click()
    Call m_modernListbox.LoadPageFromArrows(mylabel.tag)
End Sub

Private Sub myBrandLabel_Click()
    Call m_brandModernListbox.LoadPageFromArrows(myBrandLabel.tag)
End Sub


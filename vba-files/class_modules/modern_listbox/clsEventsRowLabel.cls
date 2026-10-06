VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsRowLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Events for the row labels

Private m_modernListbox As clsModernListbox
Private m_brandModernListbox As clsModernListbox
Public WithEvents mylabel As MSForms.Label
Public WithEvents myBrandLabel As MSForms.Label
Public WithEvents myTextbox As MSForms.TextBox
Public WithEvents myBrandTextbox As MSForms.TextBox

Public Property Set parentListBox(ByVal argModernListBox As clsModernListbox)
    Set m_modernListbox = argModernListBox
End Property

Public Property Set parentBrandListBox(ByVal argBrandModernListBox As clsModernListbox)
    Set m_brandModernListbox = argBrandModernListBox
End Property

Private Sub mylabel_Click()
    Call m_modernListbox.SelectRow(CLng(mylabel.tag))
End Sub

Private Sub myBrandLabel_Click()
    Call m_brandModernListbox.SelectRow(CLng(myBrandLabel.tag))
End Sub

Private Sub mylabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call m_modernListbox.DoubleClick(CLng(mylabel.tag))
End Sub

Private Sub myBrandLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call m_brandModernListbox.DoubleClick(CLng(myBrandLabel.tag))
End Sub

Private Sub myLabel_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Call m_modernListbox.SetRowHover(CLng(mylabel.tag))
End Sub

Private Sub myBrandLabel_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   Call m_brandModernListbox.SetRowHover(CLng(myBrandLabel.tag))
End Sub



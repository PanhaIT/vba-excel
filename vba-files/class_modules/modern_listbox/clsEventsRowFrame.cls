VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsRowFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Events for the row frames

Private m_modernListbox As clsModernListbox
Private m_brandModernListbox As clsModernListbox
Public WithEvents myFrame As MSForms.frame
Public WithEvents myBrandFrame As MSForms.frame

Public Property Set parentListBox(ByVal argModernListBox As clsModernListbox)
    Set m_modernListbox = argModernListBox
End Property

Public Property Set parentBrandListBox(ByVal argBrandModernListBox As clsModernListbox)
    Set m_brandModernListbox = argBrandModernListBox
End Property

Private Sub myFrame_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call m_modernListbox.SetRowHover(CLng(myFrame.tag))
End Sub

Private Sub myBrandFrame_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call m_brandModernListbox.SetRowHover(CLng(myBrandFrame.tag))
End Sub

Private Sub myFrame_Click()
    Call m_modernListbox.SelectRow(CLng(myFrame.tag))
End Sub

Private Sub myBrandFrame_Click()
    Call m_brandModernListbox.SelectRow(CLng(myBrandFrame.tag))
End Sub

Private Sub myFrame_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call m_modernListbox.DoubleClick(CLng(myFrame.tag))
End Sub

Private Sub myBrandFrame_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call m_brandModernListbox.DoubleClick(CLng(myBrandFrame.tag))
End Sub

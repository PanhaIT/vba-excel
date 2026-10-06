VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsFrameMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Events for the Main frame

Private m_modernListbox As clsModernListbox
Private m_brandModernListbox As clsModernListbox
Public WithEvents myFrame As MSForms.frame
Public WithEvents myBrandFrame As MSForms.frame

Public Property Set parentListBox(ByVal argModernListBox As clsModernListbox)
    Set m_modernListbox = argModernListBox
    Set m_brandModernListbox = argModernListBox
End Property

Private Sub myFrame_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call m_modernListbox.ClearHover
End Sub

Private Sub myBrandFrame_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call m_brandModernListbox.ClearHover
End Sub

Private Sub myFrame_Scroll(ByVal ActionX As MSForms.fmScrollAction _
                            , ByVal ActionY As MSForms.fmScrollAction _
                            , ByVal RequestDx As Single _
                            , ByVal RequestDy As Single _
                            , ByVal ActualDx As MSForms.ReturnSingle _
                            , ByVal ActualDy As MSForms.ReturnSingle)
   ' Some times the horizontal scrollbar causes issues with painting. The controls look all mixed up.
   ' This helps fixing it.
   If ActionX = fmScrollActionFocusRequest And ActualDx.Value = 0 Then
     myFrame.Repaint
   End If
End Sub

Private Sub myBrandFrame_Scroll(ByVal ActionX As MSForms.fmScrollAction _
                            , ByVal ActionY As MSForms.fmScrollAction _
                            , ByVal RequestDx As Single _
                            , ByVal RequestDy As Single _
                            , ByVal ActualDx As MSForms.ReturnSingle _
                            , ByVal ActualDy As MSForms.ReturnSingle)
   ' Some times the horizontal scrollbar causes issues with painting. The controls look all mixed up.
   ' This helps fixing it.
   If ActionX = fmScrollActionFocusRequest And ActualDx.Value = 0 Then
     myBrandFrame.Repaint
   End If
End Sub

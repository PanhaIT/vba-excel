VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsSearchLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Event for the search label. This is to simply clear the hover from the listbox row when the
' mouse is over the search label

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

Private Sub myLabel_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call m_modernListbox.ClearHover
End Sub

Private Sub myBrandLabel_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call m_brandModernListbox.ClearHover
End Sub
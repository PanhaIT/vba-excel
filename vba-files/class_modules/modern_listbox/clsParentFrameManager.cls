VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParentFrameManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is used as a workaround for a bug in VBA with causes an Automation error followed
' by Excel closing.
'
' CAUSE: If you create a frame within a frame code works fine.
' However, if
' 1. There is more than one frame on the userFrom
' 2. You create a new frame within an existing frame where the exising frame is not
'    the newest frame then Excel crashes with the automation error.

' If we use the clsModernListBox.List property this will rebuild the listbox. So when we rebuild the
' listbox we need to create a new parentFrame
'
Private m_parentFrameOriginal As MSForms.frame
Private m_parentFrame As MSForms.frame
Private m_frameMain As clsEventsFrameMain

Public Property Get parentFrame() As MSForms.frame
    Set parentFrame = m_parentFrame
End Property

Public Property Set ParentFrameOriginal(argFrame As MSForms.frame)
    Set m_parentFrameOriginal = argFrame
End Property

Public Sub InitializeFrame(settings As clsListBoxSettings, listbox As clsModernListbox)
    
    Debug.Assert Not m_parentFrameOriginal Is Nothing
    
    If Not m_parentFrame Is Nothing Then
        m_parentFrameOriginal.Parent.Controls.Remove m_parentFrame.name
    End If
    Set m_parentFrame = CreateNewParentFrame(m_parentFrameOriginal)
    
    With m_parentFrame
        .caption = ""
        .backcolor = settings.listboxBackgroundColor()
        .SpecialEffect = fmSpecialEffectFlat
#If ShowLabels = 1 Then
        .SpecialEffect = fmSpecialEffectEtched ' set to see frame
#End If
        .Visible = listbox.Visible
        Set m_frameMain = New clsEventsFrameMain
        Set m_frameMain.myFrame = m_parentFrameOriginal
        Debug.Assert Not m_parentFrameOriginal Is Nothing
        Set m_frameMain.parentListBox = listbox
        
    End With
    
End Sub


Private Function CreateNewParentFrame(myParentFrame As frame) As MSForms.frame


    With myParentFrame
        .Visible = False
       ' Set CreateNewParentFrame = .Parent.Controls.Add("Forms.Frame.1", "parentFrame", True)
        Set CreateNewParentFrame = .Parent.Controls.Add("Forms.Frame.1", myParentFrame.name & "2", True)
        CreateNewParentFrame.left = .left
        CreateNewParentFrame.top = .top
        CreateNewParentFrame.width = .width
        CreateNewParentFrame.Height = .Height
    End With

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

#If VBA7 And Win64 Then
    ' // 64bit
    Private Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#Else
    ' // 32bit
    Private Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If

Public WithEvents mForm         As MSForms.UserForm
Attribute mForm.VB_VarHelpID = -1
Public WithEvents mPage         As MSForms.MultiPage
Attribute mPage.VB_VarHelpID = -1
Public WithEvents mFrame        As MSForms.Frame
Attribute mFrame.VB_VarHelpID = -1
Public WithEvents btn           As MSForms.Label
Attribute btn.VB_VarHelpID = -1
Public WithEvents btnIcon       As MSForms.Label
Attribute btnIcon.VB_VarHelpID = -1

Public cBtn As New Collection
Private cButton As New clsBtn
Public PassiveBtnHover
Public ActiveButton As MSForms.Label

Public onButton As Boolean
Private Const IDC_HAND = 32649&

Private Function MouseCursor(CursorType As Long)
    Dim lngRet As Long
    lngRet = LoadCursorBynum(0&, CursorType)
    lngRet = SetCursor(lngRet)
End Function

Private Function MouseMoveIcon()
    Call MouseCursor(IDC_HAND)
End Function

Public Sub classButton(form As MSForms.UserForm, Optional mPage As MSForms.MultiPage, Optional mFrame As MSForms.Frame)
    Dim ctrl As control
    Dim IconCode As String
    Dim btnLabel As control

    Set mForm = form
    For Each btnLabel In mForm.Controls
        If TypeName(btnLabel) = "Label" Then
            If GetBtnType(btnLabel, 0) = "btn" Then
                ButtonDesign btnLabel

                For Each ctrl In frmElement.Controls
                    If GetBtnType(ctrl, 0) = "btn" Then
                        If GetBtnType(ctrl, 1) = GetBtnType(btnLabel, 1) Then
                            With btnLabel
                                .Picture = ctrl.Picture
                                .ForeColor = ctrl.ForeColor
                            End With
                            
                        End If
                    End If
                Next

                If btnLabel.ControlTipText <> "" Then
                    Set btnIcon = btnLabel.Parent.Controls.Add("Forms.Label.1", btnLabel.Name & "Icon")
                        With btnIcon
                        IconCode = btnLabel.ControlTipText
                            .Top = btnLabel.Top + (btnLabel.Height / 2) - (.Height / 2) + 2
                            .Left = btnLabel.Left + 6
                            .Font.Name = "myicons"
                            .Caption = ChrW("&H" & IconCode)
                            .Font.Size = 13
                            .ForeColor = btnLabel.ForeColor
                            .AutoSize = True
                            .BackStyle = fmBackStyleTransparent
                        End With
                End If

                btnLabel.ControlTipText = btnLabel.Caption

                Set cButton = New clsBtn
                Set cButton.btn = btnLabel
                Set cButton.mPage = mPage
                Set cButton.btnIcon = btnIcon
                Set cButton.mFrame = mFrame
                Set cButton.mForm = mForm
                cBtn.Add cButton

'                Set btnIcon = Nothing
            End If
        End If
    Next
End Sub

Sub ButtonDesign(ctrl As Object)
    With ctrl
        .Font.Size = 10
        .Font.Name = "Poppins Medium"
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .PicturePosition = 12
'         If InStr(1, GetBtnType(ctrl, 1), "Outline") <> 0 Then
'            .ForeColor = &H404040
'         End If
    End With
End Sub

Private Sub btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMoveIcon
    btnMouseOut
    onButton = True
    Dim ctrl As control
    For Each ctrl In frmElement.Controls
        If GetBtnType(ctrl, 0) = "btnHover" Then
            If InStr(1, GetBtnType(btn, 1), GetBtnType(ctrl, 1)) <> 0 Then
                With btn
                    .Picture = ctrl.Picture
                    .PicturePosition = fmPicturePositionCenter
                    .ForeColor = ctrl.ForeColor

                    If Not btnIcon Is Nothing Then
                        If InStr(1, btnIcon.Name, btn.Name) <> 0 Then
                            btnIcon.ForeColor = ctrl.ForeColor
                        End If
                    End If
                End With
            End If
        End If
    Next
    Set ActiveButton = btn
End Sub

Sub SetForm(form As MSForms.UserForm)
    Set mForm = form
End Sub

Private Sub mForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnMouseOut
End Sub

Sub btnMouseOut()
    Dim ctrl As control
    If onButton = True Then
        For Each ctrl In frmElement.Controls
            If GetBtnType(ctrl, 0) = "btn" Then
                If GetBtnType(ctrl, 1) = GetBtnType(ActiveButton, 1) Then
                    ActiveButton.Picture = ctrl.Picture
                    If InStr(1, GetBtnType(ctrl, 1), "Outline") <> 0 Then
                    ActiveButton.ForeColor = &H404040
                        If Not btnIcon Is Nothing Then
                        If InStr(1, btnIcon.Name, ActiveButton.Name) <> 0 Then
                            btnIcon.ForeColor = &H404040
                        End If
                        End If
                    End If
                End If
            End If
        Next
        onButton = False
    End If
End Sub

Private Sub mFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnMouseOut
End Sub

Private Sub mFrame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnMouseOut
End Sub

Private Sub mPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnMouseOut
End Sub

Function GetBtnType(ctrl As control, Index As Integer)
    On Error Resume Next
    If InStr(1, ctrl.Tag, "-") <> 0 Then
        GetBtnType = Split(ctrl.Tag, "-")(Index)
    Else
        GetBtnType = ctrl.Tag
    End If
End Function

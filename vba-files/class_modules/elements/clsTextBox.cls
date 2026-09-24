VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'//////////////////////////////////////////
'*******USERFORM CODES**********************
Public tbox As New clsTextBox
Public cBtn As New clsBtn
    
    ' Private Sub btnSaveMm_Click()
    '    tBox.TxtControl
    '    If textControl = True Then
    '        MsgBox "Thanks for feedback!"
    '    End If
    
    ' End Sub
    
    '********TEXTBOX ELEMENT***********************
    ' TxtColor ForeColor:="#a1a09f", _
    '                 EnterColor:="#F5BE57", _
    '                 TitleColor:="#b1b3b1", _
    '                 AlertColor:="#f72111", _
    '                 SuccessColor:="#22ab2b"
'     tBox.clasBox Me, "Border"
'     '********TEXTBOX ELEMENT***********************
'     cBtn.classButton Me
'*******USERFORM CODES**********************
'//////////////////////////////////////////

Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000

#If VBA7 And Win64 Then
    '// 64bit
    Private Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

    '************ Hide title bar 64bit*************
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
    '************ Hide title bar *************
#Else
    '// 32bit
    Private Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

    '************ Hide title bar 32bit*************
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    '************ Hide title bar *************
#End If

Public WithEvents myform        As MSForms.UserForm
Attribute myform.VB_VarHelpID = -1
Public WithEvents mFrame        As MSForms.Frame
Attribute mFrame.VB_VarHelpID = -1
Public WithEvents mPage         As MSForms.MultiPage
Attribute mPage.VB_VarHelpID = -1
Public WithEvents textBox       As MSForms.textBox
Attribute textBox.VB_VarHelpID = -1
Public WithEvents ComboBox      As MSForms.ComboBox
Attribute ComboBox.VB_VarHelpID = -1
Public WithEvents barBottom     As MSForms.label
Attribute barBottom.VB_VarHelpID = -1
Public WithEvents AnimationBar  As MSForms.label
Attribute AnimationBar.VB_VarHelpID = -1
Public WithEvents barTitle      As MSForms.label
Attribute barTitle.VB_VarHelpID = -1
Public WithEvents barIcon       As MSForms.label
Attribute barIcon.VB_VarHelpID = -1
Public WithEvents DropArrow     As MSForms.label
Attribute DropArrow.VB_VarHelpID = -1
Public WithEvents ControlIcon   As MSForms.label
Attribute ControlIcon.VB_VarHelpID = -1
Public WithEvents ComboFrame    As MSForms.Frame
Attribute ComboFrame.VB_VarHelpID = -1
Public WithEvents SearchBoxIcon As MSForms.label
Attribute SearchBoxIcon.VB_VarHelpID = -1

Private txtEnter As Boolean

Public txt As New clsTextBox
Public ctBox As New Collection
Private Const IDC_HAND = 32649&

Public Function HideFormTitleBar(frm As Object)
    #If VBA7 Then
        Dim lStyle As LongPtr, lFrmHandle As LongPtr
    #Else
        Dim lStyle As Long, lFrmHandle As Long
    #End If
    lFrmHandle = FindWindow("ThunderDFrame", frm.Caption)
    lStyle = GetWindowLong(lFrmHandle, GWL_STYLE)
    lStyle = lStyle And Not WS_CAPTION
    SetWindowLong lFrmHandle, GWL_STYLE, lStyle
    DrawMenuBar lFrmHandle
End Function

Private Function MouseCursor(CursorType As Long)
    Dim lngRet As Long
    lngRet = LoadCursorBynum(0&, CursorType)
    lngRet = SetCursor(lngRet)
End Function

Private Function MouseMoveIcon()
    Call MouseCursor(IDC_HAND)
End Function

Public Sub clasBox(form As MSForms.UserForm, Optional BorderType As String)
    Dim ctrl As control
    Set myform = form
    BType = BorderType
            
    For Each ctrl In myform.Controls
        If TypeName(ctrl) = "TextBox" Then
            If ctrl.Tag <> "" Then
                With ctrl
                    Call TextBoxStyle(ctrl)
                            
                    Set barBottom = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barBottom")
                    AddBarBottom ctrl, BorderType
                    
                    Set AnimationBar = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "AnimationBar")
                    AddAnimationBar ctrl, BorderType
        
                    Set barTitle = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barTitle")
                    AddBarTitle ctrl
                    
                    Set barIcon = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barIcon")
                    AddBarIcon ctrl
                    
                    Set ControlIcon = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "ControlIcon")
                    AddControlIcon ctrl
                    
                    If GetString(ctrl, 0) <> "" Then
                    .ControlTipText = GetString(ctrl, 0)
                    End If
                    
                    .ZOrder (0)
                    
                    If InStr(1, ctrl.Name, "Search") <> 0 Then
                    Set SearchBoxIcon = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "SearchBoxIcon")
                        AddSearchBoxIcon ctrl
                    End If
                End With
            
                Set txt = New clsTextBox
                Set txt.barIcon = barIcon
                Set txt.barBottom = barBottom
                Set txt.AnimationBar = AnimationBar
                Set txt.SearchBoxIcon = SearchBoxIcon
                Set txt.textBox = ctrl
                Set txt.myform = myform
                ctBox.Add txt
                
                Set barIcon = Nothing
                Set SearchBoxIcon = Nothing
            End If
        ElseIf TypeName(ctrl) = "ComboBox" Then
            If ctrl.Tag <> "" Then
                With ctrl
                    Call ComboBoxStyle(ctrl)

                    Set barBottom = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barBottom")
                        AddBarBottom ctrl, BorderType
                    
                    Set AnimationBar = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "AnimationBar")
                        AddAnimationBar ctrl, BorderType
                        
                    Set barTitle = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barTitle")
                        AddBarTitle ctrl
                    
                    Set barIcon = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barIcon")
                        AddBarIcon ctrl
                    
                    Set ControlIcon = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "ControlIcon")
                            AddControlIcon ctrl
                            
                    If GetString(ctrl, 0) <> "" Then
                        .ControlTipText = GetString(ctrl, 0)
                    End If
                        .ZOrder (0)
                    
                    Set DropArrow = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "DropArrow")
                        AddDropArrow ctrl
                    
                End With
            
                Set txt = New clsTextBox
                Set txt.ComboBox = ctrl
                Set txt.barIcon = barIcon
                Set txt.barBottom = barBottom
                Set txt.AnimationBar = AnimationBar
                Set txt.SearchBoxIcon = SearchBoxIcon
                Set txt.DropArrow = DropArrow
                Set txt.myform = myform
                ctBox.Add txt
                
                Set SearchBoxIcon = Nothing
                Set barIcon = Nothing
            End If
        ElseIf TypeName(ctrl) = "MultiPage" Then
            Set txt = New clsTextBox
            Set txt.mPage = ctrl
            Set txt.myform = myform
            ctBox.Add txt
        ElseIf TypeName(ctrl) = "Frame" Then
            Set txt = New clsTextBox
            Set txt.mFrame = ctrl
            Set txt.myform = myform
            ctBox.Add txt
        End If
    Next ctrl
End Sub

Sub TextBoxStyle(ctrl As control)
    With ctrl
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
        .SelectionMargin = False
        .MultiLine = True
        .Font.Name = FontName
        .ForeColor = fColor
        .Font.Size = FontSize
        If .Height <= 32 Then
            .Height = 32
        End If
    End With
End Sub

Sub ComboBoxStyle(ctrl As control)
    With ctrl
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
        .SelectionMargin = False
        .Font.Name = FontName
        .ForeColor = fColor
        .ShowDropButtonWhen = fmShowDropButtonWhenNever
        .Font.Size = FontSize
        If .Height <= 32 Then
            .Height = 32
        End If
    End With
End Sub

Sub AddBarBottom(ctrl As control, Optional BorderType As String)
    With barBottom
        If BorderType = "Border" Then
            If ctrl.Height > 32 Then
               .Height = ctrl.Height '+ 2
            Else
               .Height = 32
                ctrl.Height = 32
            End If
            .Top = ctrl.Top - 2
        Else
            
            .Height = 1
            .Top = ctrl.Top + ctrl.Height - 3
        End If
         
        .Width = ctrl.Width
        .BorderStyle = fmBorderStyleSingle
        .Left = ctrl.Left
        .BackStyle = fmBackStyleTransparent
        .BorderColor = tColor
        .ZOrder (0)
    End With
End Sub

Sub AddAnimationBar(ctrl As control, Optional BorderType As String)
    With AnimationBar
        .Height = 1
        .BackColor = eColor
        .Top = ctrl.Top + ctrl.Height - 3
        .Width = 0
        .BorderStyle = fmBorderStyleNone
        .Left = ctrl.Left
        .ZOrder (0)
    End With
End Sub

Sub AddBarTitle(ctrl As control)
    With barTitle
        If GetString(ctrl, 1) <> "" Then
        .Caption = Space(1) & GetString(ctrl, 1)
        Else
        .Caption = Space(1) & GetString(ctrl, 0)
        End If
        .ForeColor = tColor
        .Font.Name = FontName
        .Font.Size = FontSize
        .BackStyle = fmBackStyleTransparent
        .Top = ctrl.Top + 6
        .ZOrder (0)
        .AutoSize = True
        .WordWrap = False
        If ctrl.ControlTipText <> "" Then
        .Left = ctrl.Left + 20
        Else
        If BType = "Border" Then
            With ctrl
                .Left = .Left + 6
                .Width = .Width - 8
            End With
        End If
        .Left = ctrl.Left
        End If
    '.BorderStyle = fmBorderStyleSingle
    End With
End Sub

Sub AddBarIcon(ctrl As control)
    With barIcon
        .Font.Name = IconFont
        .Font.Size = 14
        .Width = 18
        .Height = 18
        .ForeColor = tColor
        .Top = ctrl.Top + 6
        .Left = ctrl.Left + 4
        .BackStyle = fmBackStyleTransparent
         If ctrl.ControlTipText <> "" Then
            .Caption = ChrW("&H" & ctrl.ControlTipText)
             With ctrl
                .Left = .Left + 20
                .Width = .Width - 22
            End With
          End If
        .ZOrder (0)
    End With
End Sub

Sub AddDropArrow(ctrl As control)
    With DropArrow
        .Font.Name = "Webdings"
        .Caption = "6"
        .Height = 18
        .Width = 18
        .Left = ctrl.Width + ctrl.Left - .Width + 6
        .BackStyle = fmBackStyleTransparent
        .ForeColor = tColor
        .BorderStyle = fmBorderStyleNone
        .Top = ctrl.Top + 4
        .Font.Size = FontSize
        .ZOrder (0)
    End With
End Sub

Sub AddControlIcon(ctrl As control)
    With ControlIcon
        .Font.Name = "myicons"
        .Caption = ChrW("&HE012")
        .AutoSize = True
        .WordWrap = False
        .Left = ctrl.Width + ctrl.Left - 22
        .BackStyle = fmBackStyleTransparent
        .ForeColor = aColor
        .BorderStyle = fmBorderStyleNone
        .Top = ctrl.Top + 6
        .Font.Size = FontSize
        .Visible = False
        .ZOrder (0)
    End With
End Sub

Public Sub AddSearchBoxIcon(ctrl As control)
    With SearchBoxIcon
        .Font.Name = "myicons"
        .Caption = ChrW("&HE012")
        .AutoSize = True
        .WordWrap = False
        .Left = ctrl.Width + ctrl.Left - 16
        .BackStyle = fmBackStyleTransparent
        .ForeColor = aColor
        .BorderStyle = fmBorderStyleNone
        .Top = ctrl.Top + 6
        .Font.Size = FontSize - 1
        .Visible = False
        .ControlTipText = "Reset"
        .ZOrder (0)
    End With
End Sub

Sub SetForm(form As MSForms.UserForm)
    Set txt = New clsTextBox
    Set txt.myform = form
    ctBox.Add txt
End Sub

Private Sub barIcon_Click()
    With barIcon
        If .Caption = ChrW("&H005C") Then
           If textBox.Locked = False Then
                TextBoxEnter textBox
                Call DatePicker.ChooseDate(myform.Controls(textBox.Name))
           End If
        End If
    End With
End Sub

Private Sub barIcon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     With barIcon
        If .Caption = ChrW("&H005C") Then
            MouseMoveIcon
        End If
    End With
End Sub

Private Sub class_initialize()
    textControl = True
    txtEnter = True
End Sub

Private Sub Class_Terminate()
    Set ctBox = Nothing
    Set txt = Nothing
End Sub

Private Sub mFrame_Click()
    BoxExit
End Sub

Private Sub mPage_Change()
    textControl = True
    BoxExit
End Sub

Private Sub mPage_Click(ByVal Index As Long)
    BoxExit
End Sub

Private Sub mPage_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    BoxExit
End Sub

Private Sub myform_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    BoxExit
End Sub

Private Sub SearchBoxIcon_Click()
    textBox = ""
    TextBoxEnter textBox
    textBox.SetFocus
End Sub

Private Sub SearchBoxIcon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMoveIcon
End Sub

Private Sub TextBox_Change()
    If txtEnter = True Then
        BoxExit
    End If
    If InStr(1, textBox.Name, "Search") <> 0 And textBox.Name <> "SearchBox" Then
        If Len(textBox) > 0 Then
            SearchBoxIcon.Visible = True
        Else
            SearchBoxIcon.Visible = False
        End If
    End If
End Sub

Private Sub Textbox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxEnter textBox
End Sub

Private Sub Textbox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then TextBoxEnter textBox
End Sub

Sub TextBoxEnter(Optional ctrl As control)
    If ActiveBox <> ctrl.Name Then BoxExit
    If textBox.Tag <> "" Then
        ControlIconVisible textBox, tColor, False, "Success"
        BarTitleForeColor textBox, eColor, False
        SearchBoxColor textBox, eColor
        BarAnimation textBox, True, eColor
    End If
    txtEnter = False
    ActiveBox = ctrl.Name
End Sub

Sub BoxExit(Optional mPage As MSForms.MultiPage, Optional PageIndex As Integer)
    Dim control As control
    If Not mPage Is Nothing Then
        Set ControlArea = myform.Controls(mPage.Name).pages(PageIndex)
    Else
        Set ControlArea = myform
    End If

    For Each control In ControlArea.Controls
    If (TypeName(control) = "TextBox" Or TypeName(control) = "ComboBox") Then
        If control.Tag <> "" Then
            If Len(control) = 0 Then
                If textControl = True Then

                    BarAnimation control, False, tColor
                    BarTitleForeColor control, tColor, True
                    ControlIconVisible control, tColor, False
                    SearchBoxColor control, tColor, True
                    
                    If TypeName(control) = "ComboBox" Then
                        DropArrowForeColor control, tColor
                    End If
                    
                    Else
                        If InStr(1, control.Tag, "*") <> 0 Then
                            BarAnimation control, False, aColor
                            BarBottomBorderColor control, aColor
                            BarTitleForeColor control, aColor, True
                            ControlIconVisible control, aColor, True, "Alert"
                            If TypeName(control) = "ComboBox" Then
                                DropArrowForeColor control, aColor
                            End If
                        Else
                            BarAnimation control, False, sColor
                            ControlIconVisible control, sColor, True, "Success"
                            BarTitleForeColor control, sColor, True
                            SearchBoxColor control, tColor, True
                            
                            If TypeName(control) = "ComboBox" Then
                                DropArrowForeColor control, sColor
                            End If
                        End If
                    End If
                Else
                    If textControl = True Then
                        BarAnimation control, False, tColor
                        ControlIconVisible control, tColor, False, "Success"
                        BarTitleForeColor control, tColor, False
                        SearchBoxColor control, tColor
                        
                        If TypeName(control) = "ComboBox" Then
                            DropArrowForeColor control, tColor
                        End If
                    Else
                        BarAnimation control, False, sColor
                        ControlIconVisible control, sColor, True, "Success"
                        BarTitleForeColor control, sColor, False
                        SearchBoxColor control, tColor, False
                        
                        If TypeName(control) = "ComboBox" Then
                            DropArrowForeColor control, sColor
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub ComboBoxEnter(Optional ctrl As control)
    If Not ctrl Is Nothing Then
        If ActiveBox <> Empty Then
            If ActiveBox <> ctrl.Name And _
               ActiveBox <> myform.Controls(ctrl.Name & "DropArrow").Name Then
               BoxExit
            End If
        End If
    End If
    
    If ComboBox.Tag <> "" Then
        ControlIconVisible ComboBox, tColor, False, "Success"
        BarTitleForeColor ComboBox, eColor, False
        BarAnimation ComboBox, True, eColor
        DropArrowForeColor ComboBox, eColor
     End If
     txtEnter = False
     ActiveBox = ctrl.Name
End Sub

Private Sub ComboBox_Change()
    If txtEnter = True Then ComboBoxEnter ComboBox
End Sub

Private Sub ComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
     If KeyCode = 9 Then ComboBoxEnter ComboBox
End Sub

Private Sub ComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ComboBoxEnter ComboBox
End Sub

Private Sub DropArrow_Click()
    Dim comboName As String
    Dim ctrl As control

    comboName = Mid(DropArrow.Name, 1, Len(DropArrow.Name) - 9)
    ' comboName = Mid(DropArrow.Name, 1, Len(DropArrow.Name) - 9)
    If ActiveBox <> comboName And ActiveBox <> DropArrow.Name Then BoxExit

    For Each ctrl In myform.Controls
        If TypeName(ctrl) = "ComboBox" Then
            comboName = Mid(DropArrow.Name, 1, Len(DropArrow.Name) - 9)
            'comboName = Mid(DropArrow.Name, 1, Len(DropArrow.Name) - 9)
            If comboName = ctrl.Name Then
                With myform.Controls(comboName)
                    .DropDown
                    
                    ControlIconVisible myform.Controls(comboName), tColor, False, "Success"
                    BarBottomBorderColor myform.Controls(comboName), eColor
                    BarAnimation myform.Controls(comboName), True, eColor
                    BarTitleForeColor myform.Controls(comboName), eColor, False
                    BarIconForeColor myform.Controls(comboName), eColor
                End With

                With DropArrow
                    .ForeColor = eColor
                End With
            End If
        End If
    Next ctrl
    ActiveBox = DropArrow.Name
End Sub

Private Sub DropArrow_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMoveIcon
End Sub

Private Sub myForm_Click()
    BoxExit
End Sub

Sub ControlIconVisible(ctrl As control, ForeColor As Variant, visibilty As Boolean, Optional IconType As String)
    With myform.Controls(ctrl.Name & "ControlIcon")
        If visibilty = True Then
            .Visible = True
            If IconType = "Alert" Then
                .Caption = ChrW("&HE0B2")
            Else
                .Caption = ChrW("&HE03B")
                .ForeColor = ForeColor
            End If
        Else
            .Visible = False
        End If
        .ForeColor = ForeColor
    
        If InStr(1, ctrl.Name, "Search") <> 0 Then
            .Visible = False
        End If
    End With
End Sub

Sub BarBottomBorderColor(ctrl As control, BorderColor As Variant)
    With myform.Controls(ctrl.Name & "BarBottom")
      
      .BorderColor = BorderColor
    End With
End Sub

Sub BarTitleForeColor(ctrl As control, ForeColor As Variant, TitlePosition As Boolean)
    With myform.Controls(ctrl.Name & "BarTitle")
        .ForeColor = ForeColor
        If TitlePosition = True Then
            .Top = ctrl.Top + 4 'text in input box position
            .Left = ctrl.Left
            .Font.Size = FontSize
            If GetString(ctrl, 1) <> "" Then
                .Caption = Space(1) & GetString(ctrl, 1)
            Else
                .Caption = Space(1) & GetString(ctrl, 0)
            End If
        Else
            .Font.Size = FontSize - 1
            .BackColor = myform.BackColor
            .Top = ctrl.Top - 12
            .Caption = GetString(ctrl, 0)

            If BType = "Line" Then
                .Left = myform.Controls(ctrl.Name & "barBottom").Left - 1 'Textbox.Left + 4
                .BackStyle = fmBackStyleTransparent
            Else
                .BackStyle = fmBackStyleOpaque
                .Left = myform.Controls(ctrl.Name & "barBottom").Left + 4
            End If
            
        End If
        BarIconForeColor ctrl, ForeColor
        BarBottomBorderColor ctrl, ForeColor
    End With
End Sub

Sub SearchBoxColor(ctrl As control, ForeColor As Variant, Optional TitlePosition As Boolean)
    If InStr(1, ctrl.Name, "Search") <> 0 Then
         BarTitleForeColor ctrl, ForeColor, TitlePosition
         BarIconForeColor ctrl, ForeColor
         BarBottomBorderColor ctrl, ForeColor
    End If
End Sub

Sub BarIconForeColor(ctrl As control, ForeColor As Variant)
    With myform.Controls(ctrl.Name & "BarIcon")
      .ForeColor = ForeColor
    End With
End Sub

Sub BarAnimation(ctrl As control, BarState As Boolean, BackColor As Variant)
    Dim spd As Double
    barBottomElement = myform.Controls(ctrl.Name & "barBottom")
    With myform.Controls(ctrl.Name & "AnimationBar")
        If BarState = True Then
            spd = barBottom.Width / (barBottom.Width / (barBottom.Width * 1.1)) / 250
            Do While .Width < barBottom.Width
            DoEvents
                .Width = .Width + spd
            Loop
            .BackColor = BackColor
            .Width = barBottom.Width
        Else
'             spd = barBottom.Width / (barBottom.Width / (barBottom.Width * 1.1)) / 100
             Do While .Width >= 5
             DoEvents
                 If .Width > 0 Then .Width = .Width - 2
             Loop
            .Width = 0

        End If
    End With
End Sub

Sub DropArrowForeColor(ctrl As control, ForeColor As Variant)
    With myform.Controls(ctrl.Name & "DropArrow")
      .ForeColor = ForeColor
    End With
End Sub

Public Sub TxtControl(Optional mPage As MSForms.MultiPage, Optional PageIndex As Integer)
    textControl = True
    Dim ctrl As control
    
    If Not mPage Is Nothing Then
     Set ControlArea = myform.Controls(mPage.Name).pages(PageIndex)
         For Each ctrl In ControlArea.Controls
            If (TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox") Then
                If InStr(1, ctrl.Tag, "*") <> 0 And ctrl.Visible = True Then
                    If ctrl = "" Then
                        textControl = False
                    End If
                End If
            End If
        Next ctrl
         BoxExit myform.Controls(mPage.Name), PageIndex
    Else
        For Each ctrl In myform.Controls
            If (TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox") Then
                If InStr(1, ctrl.Tag, "*") <> 0 And ctrl.Visible = True Then
                    If ctrl = "" Then
                        textControl = False
                    End If
                End If
            End If
        Next ctrl
        BoxExit
    End If
End Sub
    
Function GetString(ctrl As control, Index As Integer)
    If InStr(1, ctrl.Tag, "-") <> 0 Then
        GetString = Split(ctrl.Tag, "-")(Index)
    Else
        GetString = ctrl.Tag
    End If
End Function



Attribute VB_Name = "mdlMinimizeMaximize"
'**********USERFORM CODES**********************

'Private Sub UserForm_Activate()
'    MakeFormResizable
'    If FormLoad = False Then _
'        SetStandAloneForm Me, "Form"
'    FormLoad = True
'End Sub

'**********USERFORM CODES**********************


#If VBA7 Then
    
    Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

    Public Declare PtrSafe Function SetWindowPos Lib "user32" _
                                          (ByVal hWnd As Long, _
                                           ByVal hWndInsertAfter As Long, _
                                           ByVal X As Long, _
                                           ByVal Y As Long, _
                                           ByVal cx As Long, _
                                           ByVal cy As Long, _
                                           ByVal wFlags As Long) As Long

    Public Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As Long
    Public Declare PtrSafe Function SendMessage Lib "user32" _
                                         Alias "SendMessageA" _
                                         (ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          ByVal wParam As Long, _
                                          lParam As Any) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
                                     (ByVal hWnd As Long) As Long
                                     
    Public Declare PtrSafe Function GetFocus Lib "user32" () As Long
#Else
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long

    Public Declare Function SetWindowPos Lib "user32" _
                                          (ByVal hWnd As Long, _
                                           ByVal hWndInsertAfter As Long, _
                                           ByVal X As Long, _
                                           ByVal Y As Long, _
                                           ByVal cx As Long, _
                                           ByVal cy As Long, _
                                           ByVal wFlags As Long) As Long

    Public Declare Function GetActiveWindow Lib "user32.dll" () As Long
    Public Declare Function SendMessage Lib "user32" _
                                         Alias "SendMessageA" _
                                         (ByVal hWnd As Long, _
                                          ByVal wMsg As Long, _
                                          ByVal wParam As Long, _
                                          lParam As Any) As Long
    Public Declare Function DrawMenuBar Lib "user32" _
                                     (ByVal hWnd As Long) As Long

    Private Declare Function GetFocus Lib "user32" () As Long
#End If

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const WS_EX_APPWINDOW = &H40000
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const SWP_FRAMECHANGED = &H20
Public Const WM_SETICON = &H80
Public Const ICON_SMALL = 0&
Public Const ICON_BIG = 1&

Public FormLoad As Boolean
Public ActiveNotLabel, ActiveMarkRead

Public Sub MakeFormResizable()
    Dim lStyle As Long
    Dim hWnd As Long
    Dim RetVal
    
    Const WS_THICKFRAME = &H40000
    Const GWL_STYLE As Long = (-16)
    
    hWnd = GetActiveWindow
    lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME
    RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)
    If RetVal = 0 Then MsgBox "Unable to make UserForm Resizable."
End Sub

Public Function SetStandAloneForm(form As Object, Optional PicName As String = "Form")
    Const GWL_STYLE As Long = -16
    Const GWL_EXSTYLE As Long = -20
    Const WS_CAPTION As Long = &HC00000
    Const WS_MINIMIZEBOX As Long = &H20000
    Const WS_MAXIMIZEBOX As Long = &H10000
    Const WS_POPUP As Long = &H80000000
    Const WS_VISIBLE As Long = &H10000000
    Const WS_EX_DLGMODALFRAME As Long = &H1
    Const WS_EX_APPWINDOW As Long = &H40000
    Const SW_SHOW As Long = 5

    Dim hWnd As Long
    Dim CurrentStyle As Long
    Dim NewStyle As Long
    Dim lngRet As Long
    Dim hIcon As Long
    Dim img As MSForms.Image

    If Val(Application.Version) < 9 Then
        hWnd = FindWindow("ThunderXFrame", form.Caption)  'XL97
    Else
        hWnd = FindWindow("ThunderDFrame", form.Caption)  '>XL97
    End If

    CurrentStyle = GetWindowLong(hWnd, GWL_STYLE)
    NewStyle = CurrentStyle Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
    NewStyle = NewStyle And Not WS_VISIBLE And Not WS_POPUP
    Call SetWindowLong(hWnd, GWL_STYLE, NewStyle)

    CurrentStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    NewStyle = CurrentStyle Or WS_EX_APPWINDOW
    Call SetWindowLong(hWnd, GWL_EXSTYLE, NewStyle)
    Call ShowWindow(hWnd, SW_SHOW)

    Set img = form.Controls.Add("Forms.Image.1", "img")
        With img
            .Visible = False
            .Picture = frmElement.formicon.Picture
        End With
    hIcon = form.img.Picture.Handle

    hWnd = FindWindow(vbNullString, form.Caption)
    lngRet = SendMessage(hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
    lngRet = SendMessage(hWnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    lngRet = DrawMenuBar(hWnd)
End Function



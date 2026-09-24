VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'********MOUSE MOVE ICON CODES*************************************************************************************************************
#If VBA7 And Win64 Then
    ' // 64bit
    Private Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#Else
    ' // 32bit
    Private Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
#End If
'    Private Function MouseMoveIcon()
'        Dim lngRet As Long
'        lngRet = LoadCursorBynum(0&, 32649&)
'        lngRet = SetCursor(lngRet)
'    End Function
'********MOUSE MOVE ICON CODES*************************************************************************************************************

Public WithEvents mFrame             As MSForms.Frame
Public WithEvents pageFrame          As MSForms.Frame
Public WithEvents btnFirstPage       As MSForms.label
Public WithEvents btnPreviewPage     As MSForms.label
Public WithEvents btnNextPage        As MSForms.label
Public WithEvents btnLastPage        As MSForms.label

Public WithEvents FirstRec           As MSForms.label
Public WithEvents LastRec            As MSForms.label
Public WithEvents TotalRec           As MSForms.label
Public WithEvents hyphen             As MSForms.label
Public WithEvents Slash              As MSForms.label
Public WithEvents firstPageBack      As MSForms.label
Public lblLeft As Integer

' Private Function MouseMoveIcon()
'     Dim lngRet As Long
'     lngRet = LoadCursorBynum(0&, 32649&)
'     lngRet = SetCursor(lngRet)
' End Function

Sub setPageElement(muframe As MSForms.Frame, Optional paFrame As MSForms.Frame, Optional TPageLabel As MSForms.label)
    Set mFrame = muframe
    Set pFrame = paFrame
    Set TotalPageLabel = TPageLabel
End Sub

Sub AddPageElement(mFrame As MSForms.Frame, Optional pFrame As MSForms.Frame, Optional TPageLabel As MSForms.label)
    If Not pFrame Is Nothing Then _
    Set pageFrame = pFrame
    Set TotalPageLabel = TPageLabel

    With pageFrame
        .Clear
        .Caption = ""
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone 'fmBackStyleTransparent
        '.BackColor = SortIconHead 'NavFrameColor
        '.BorderColor = SortIconHead
        .Top = 0 'mFrame.Height + mFrame.Top - 20 '20
        .Left = 0 'mFrame.Width + mFrame.Left - pageFrame.Width - 20 ' -20
        .Visible = True
        .Width = 280
        .Height = 40
        
        totalpage = Application.WorksheetFunction.RoundUp(TotalRecVal / ListRows, 0)
        With TotalPageLabel
            .Left = 10
            .Font.Name = "Roboto" 'Medium
            .Font.Size = 11
            .ForeColor = vbGrayText
            '3rd page out of 5 total pages
             '.Caption = totalpage & " Pages " & PageNumber & ".Page"
            .Caption = "Page." & PageNumber & " of " & totalpage & " total Pages"
            .Top = mFrame.Height + mFrame.Top + 20
            .AutoSize = True
            .WordWrap = False
            .BackStyle = fmBackStyleTransparent
            '.BorderStyle = fmBorderStyleSingle
            .ZOrder 1
        End With
        
        lblLeft = 30
        
       Set FirstRec = .Controls.Add("Forms.Label.1", "FirstRec")
            LabelDesign FirstRec, FirstRecVal
        
        Set hyphen = .Controls.Add("Forms.Label.1", "hyphen")
            LabelDesign hyphen, , "-"
            
        Set LastRec = .Controls.Add("Forms.Label.1", "LastRec")
            LabelDesign LastRec, LastRecVal
                With LastRec
                    If rowCount > LastRecVal Then
                        .Caption = LastRecVal
                    Else
                        If rowCount > 1 Then
                        .Caption = rowCount
                        Else
                        .Caption = 0
                        End If
                        
                    End If
                End With
        Set Slash = .Controls.Add("Forms.Label.1", "Slash")
        LabelDesign Slash, , " of "
                   
        Set TotalRec = .Controls.Add("Forms.Label.1", "TotalRec")
        LabelDesign TotalRec, TotalRecVal
        
        Set btnFirstPage = .Controls.Add("Forms.Label.1", "btnFirstPage")
            ButtonDesign btnFirstPage, "E0BE", 150, "First Page"

        Set btnPreviewPage = .Controls.Add("Forms.Label.1", "btnPreviewPage")
            ButtonDesign btnPreviewPage, "E04D", 180, "Previous Page"

        Set btnNextPage = .Controls.Add("Forms.Label.1", "btnNextPage")
            ButtonDesign btnNextPage, "E04C", 210, "Next Page"
            
        Set btnLastPage = .Controls.Add("Forms.Label.1", "btnLastPage")
            ButtonDesign btnLastPage, "E0BF", 240, "Last Page"
        
        Set firstPageBack = .Controls.Add("Forms.Label.1", "firstPageBack")
        With firstPageBack
            .Height = 25
            .Width = 30
            .Left = 140
            .Top = 3
            .BackStyle = fmBorderStyleSingle
            .BackColor = BackLabelMenu 'RowBackBackColor ' DeleteIconActive 'MenuBorder
            .ZOrder 1
            .Visible = False
        End With
    End With
End Sub

Sub LabelDesign(lbl As control, Optional Caption As Integer, Optional symbol As String)
    With lbl
        .Font.Name = "Roboto" 'Medium
        .Font.Size = 11
        .ForeColor = PaginationHover
         If Caption > 0 Then
            .Caption = Caption
         Else
            .Caption = symbol
         End If
        .Left = lblLeft
        .AutoSize = True
        .BackStyle = fmBackStyleTransparent
        .Top = 10
        lblLeft = lblLeft + .Width + 4

    End With
End Sub

Sub ButtonDesign(btn As control, IconCode As String, btnLeft As Integer, TipTex As String)
    With btn
        .Font.Name = "myicons"
        .Font.Size = 11
        .ForeColor = ActivePagination 'RGB(169, 174, 177) #a9aeb1
        .Caption = ChrW("&H" & IconCode)
        .ControlTipText = TipTex
        .Left = btnLeft
        .BackStyle = fmBackStyleTransparent
        .Top = 10
    End With
End Sub

Private Sub btnFirstPage_Click()
    PageNumber = 1
    FirstRecVal = 1
    LastRecVal = ListRows
    ButtonClick
End Sub

Private Sub btnPreviewPage_Click()
   If FirstRecVal > 1 Then
        FirstRecVal = FirstRecVal - ListRows
        LastRecVal = FirstRecVal + ListRows - 1
        If PageNumber > 1 Then PageNumber = PageNumber - 1
        If FirstRecVal <= 0 Then FirstRecVal = 1
        ButtonClick
    End If
End Sub

Private Sub btnNextPage_Click()
    If LastRecVal < rowCount Then
        FirstRecVal = LastRecVal + 1
        LastRecVal = LastRecVal + ListRows
        If PageNumber < totalpage Then PageNumber = PageNumber + 1
        If LastRecVal > rowCount Then
            LastRecVal = rowCount
'            FirstRecVal = LastRecVal - ListRows
        End If
        ButtonClick
    End If
End Sub

Private Sub btnLastPage_Click()
    If LastRecVal < TotalRecVal Then
        LastRecVal = TotalRecVal
        FirstRecVal = (Application.WorksheetFunction.RoundDown((rowCount / ListRows), 0) * ListRows) + 1
        PageNumber = totalpage
        ButtonClick
    End If
End Sub

Sub ButtonClick()
'    AddPageElement
    With frmMain
        If frmMain.mp_menu.Value = 1 Then '##########
            product.TableDataList .txtSearchPerson,True, "ID", "ASC"
        End If
    End With
End Sub

Sub PagingMouseOut()
    Dim ctrl As control
    For Each ctrl In pageFrame.Controls
        btnFirstPage.ForeColor = ActivePagination 'RGB(169, 174, 177) #a9aeb1
        btnPreviewPage.ForeColor = ActivePagination 'RGB(169, 174, 177) #a9aeb1
        btnNextPage.ForeColor = ActivePagination 'RGB(169, 174, 177) #a9aeb1
        btnLastPage.ForeColor = ActivePagination 'RGB(169, 174, 177) #a9aeb1
        If ctrl.Name = "firstPageBack" Then
            With pageFrame
                .Controls("firstPageBack").Visible = False
            End With
        End If
    Next
End Sub

Private Sub pageFrame_Click()
    PagingMouseOut
End Sub

Private Sub pageFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    PagingMouseOut
End Sub

Private Sub btnFirstPage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    MouseMoveIcon
    PagingMouseOut
    btnFirstPage.ForeColor = PaginationHover 'RGB(113, 118, 122) #71767a
    With pageFrame
        .Controls("firstPageBack").Visible = True
        .Controls("firstPageBack").Left = btnFirstPage.Left - 10
    End With
End Sub

Private Sub btnLastPage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    MouseMoveIcon
    PagingMouseOut
    btnLastPage.ForeColor = PaginationHover 'RGB(113, 118, 122) #71767a
    With pageFrame
        .Controls("firstPageBack").Visible = True
        .Controls("firstPageBack").Left = btnLastPage.Left - 10
    End With
End Sub

Private Sub btnNextPage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    MouseMoveIcon
    PagingMouseOut
    btnNextPage.ForeColor = PaginationHover 'RGB(113, 118, 122) #71767a
    With pageFrame
        .Controls("firstPageBack").Visible = True
        .Controls("firstPageBack").Left = btnNextPage.Left - 10
    End With
End Sub

Private Sub btnPreviewPage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    MouseMoveIcon
    PagingMouseOut
    btnPreviewPage.ForeColor = PaginationHover 'RGB(113, 118, 122) #71767a
    With pageFrame
        .Controls("firstPageBack").Visible = True
        .Controls("firstPageBack").Left = btnPreviewPage.Left - 10
    End With
End Sub

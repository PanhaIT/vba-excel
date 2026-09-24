VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public WithEvents mForm            As MSForms.UserForm
Attribute mForm.VB_VarHelpID = -1
Public WithEvents mFrame           As MSForms.Frame
Attribute mFrame.VB_VarHelpID = -1

Public WithEvents PreviewButton    As MSForms.Label
Attribute PreviewButton.VB_VarHelpID = -1
Public WithEvents MonthLabel       As MSForms.Label
Attribute MonthLabel.VB_VarHelpID = -1
Public WithEvents YearLabel        As MSForms.Label
Attribute YearLabel.VB_VarHelpID = -1
Public WithEvents NextButton       As MSForms.Label
Attribute NextButton.VB_VarHelpID = -1

Public WithEvents DaysLabel        As MSForms.Label
Attribute DaysLabel.VB_VarHelpID = -1
Public WithEvents DateLabel        As MSForms.Label
Attribute DateLabel.VB_VarHelpID = -1
Public WithEvents DateCircle       As MSForms.Label
Attribute DateCircle.VB_VarHelpID = -1

Public DaysLeft         As Integer
Public LabelTop         As Integer
Public LabelLeft        As Integer

Public this             As New clsDate
Public thisCol          As New Collection

Private Const IDC_HAND = 32649

Public Sub MouseMoveIcon()
    Dim handCursor As Long
    handCursor = LoadCursor(0, IDC_HAND)
    
    SetCursor handCursor
End Sub

Public Sub CalendarCreate(cYear As Integer, cMonth As Integer, Optional cDay As Date)
    Dim firstDay            As Date
    Dim DateNumber          As Integer
    Dim k                   As Integer

    If cDay = "00:00:00" Then cDay = Date

    Set mForm = DatePicker
    Set mFrame = DatePicker.FrameDatePicker

    With mFrame
        .Clear
        .Caption = ""
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone

        '//Direction Button
        LabelTop = 5
        Set PreviewButton = .Controls.Add("Forms.Label.1", "PreviewButton")
        IconDesign PreviewButton, "E04D", 14
        With PreviewButton
            .Left = 15
        End With

        Set NextButton = .Controls.Add("Forms.Label.1", "NextButton")
        IconDesign NextButton, "E04C", 14
            With NextButton
            .Left = mFrame.Width - .Width - 20
        End With

        '//Month and Year Label
        Set MonthLabel = .Controls.Add("Forms.Label.1", "MonthLabel")
        LabelDesign MonthLabel, Format(MonthName(cMonth, True), "mmm")
        With MonthLabel
            .Left = mFrame.Width / 2.74
            .Font.Size = 11
            .Tag = cMonth
        End With
                
        Set YearLabel = .Controls.Add("Forms.Label.1", "YearLabel")
        LabelDesign YearLabel, cYear
        With YearLabel
            .Left = mFrame.Width / 1.86
            .Font.Size = 11
        End With

        LabelTop = 30: LabelLeft = 15
        
        For i = FirstWeek To FirstWeek + 6
            Set DaysLabel = .Controls.Add("Forms.Label.1", "DaysLabel" & i)
            LabelDesign DaysLabel, Format(i, "ddd")
        Next
        
        Set DateCircle = .Controls.Add("Forms.Label.1", "DateCircle")
        IconDesign DateCircle, "E125", 22
        With DateCircle
            .Visible = False
            .ForeColor = RGB(243, 170, 96)
            .ZOrder 1
        End With
            
        '//G�nler Ekleniyor
        dt = DateSerial(cYear, cMonth, 1)
        firstDate = Weekday(dt, FirstWeek)
        firstDay = DateSerial(Year(dt), Month(dt), 1 - (firstDate - 1))
        
        k = 1
        LabelTop = 50: LabelLeft = 20
        For i = firstDay To firstDay + 41
            
            DateNumber = Format(Day(i), "#0")
            If Len(DateNumber) = 1 Then DateNumber = " " & DateNumber
            
            Set DateLabel = .Controls.Add("Forms.Label.1", "DateLabel" & i)
            LabelDesign DateLabel, DateNumber
            With DateLabel

                .Tag = Format(i, DateType)
                .Font.Name = "Nunito Medium"
                
                '//Sadece bu ay�n g�nleri
                If Month(i) = Month(dt) Then
                    .Enabled = True
'                      If i > Date + 30 Then .Enabled = False
                
                Else
                    .Enabled = False
                End If
                
                '//tarih bo� geldi�inde veya bug�n�n tarihi oldu�unda
                If i = cDay Then
                    .ForeColor = vbWhite
                    Set ActiveDate = DateLabel
                    OldForeColor = vbGrayText
                    DateCirclePosition DateLabel
                End If
                
                If i = Date And i = cDay Then
                    .ForeColor = vbWhite
                    Set ActiveDate = DateLabel
                    OldForeColor = RGB(243, 170, 96)
                    DateCirclePosition DateLabel
                End If
                
                '//E�er tarih bug�n ise ve se�ili tarih de�ilse
                If i = Date And i <> cDay Then
                    .ForeColor = RGB(243, 170, 96)
                End If
                
                '//E�er g�nler haftasonuna denk geliyorsa
                If Weekday(i, FirstWeek) = 6 Or Weekday(i, FirstWeek) = 7 Then
                    .ForeColor = vbRed
                    If i = cDay Then
                        .ForeColor = vbWhite
                        Set ActiveDate = DateLabel
                        OldForeColor = vbRed
                    End If
                End If
                
            End With
            
            If k Mod 7 = 0 Then
                LabelLeft = 20
                LabelTop = LabelTop + 25
            End If
            
            '//Nesneleri koleksiyona ekleyerek hepsinin tek bir nesne gibi davranmas�n� sa�l�yoruz
            Set this = New clsDate
            Set this.DateLabel = DateLabel
            Set this.DateCircle = DateCircle
            Set this.mFrame = mFrame
            thisCol.Add this
            
            k = k + 1
        Next    
    End With
End Sub

Public Sub LabelDesign(ctrl As MSForms.Label, Caption)
    With ctrl
        .Left = LabelLeft
        .AutoSize = True
        .WordWrap = False
        .Font.Name = "Nunito ExtraBold"
        .Font.Size = 10
        .ForeColor = vbGrayText
        .Caption = Caption
        .BackStyle = fmBackStyleTransparent
        .Top = LabelTop
        LabelLeft = LabelLeft + mFrame.Width / 7
    End With
End Sub

Public Sub IconDesign(ctrl As MSForms.Label, IconCode As String, Optional FontSize As Integer = 18)
    With ctrl
        .Font.Name = "myicons"
        .Font.Size = FontSize
        .ForeColor = vbGrayText
        .Caption = ChrW("&H" & IconCode)
        .BackStyle = fmBackStyleTransparent
        .Top = LabelTop
        .AutoSize = True
        .WordWrap = False
    End With
End Sub

Private Sub DateLabel_Click()
    If Not ActiveDate Is Nothing Then
        ActiveDate.ForeColor = vbGrayText
        ActiveDate.ForeColor = OldForeColor
    End If
    
    With DateLabel
         Set ActiveDate = DateLabel
         OldForeColor = DateLabel.ForeColor
         
         DatePicker.CurrentDate = .Tag
        .ForeColor = vbWhite
    
    SelectDay .Tag
    End With
End Sub

Public Sub DateCirclePosition(ctrl As control)
    With DateCircle
        .Top = ctrl.Top + (ctrl.Height / 2) - (.Height / 2)
        .Left = ctrl.Left + (ctrl.Width / 2) - (.Width / 2)
        .Visible = True
    End With
End Sub

Public Sub SelectDay(SelectedDate As Date)
    With DatePicker
        .lblSelectedDay = Format(Day(SelectedDate), "0#")
        .lblSelectedDayName = Format(SelectedDate, "dddd")
        .lblSelectedMonth = MonthName(Month(SelectedDate))
        .lblSelectedYear = Year(SelectedDate)
    End With
End Sub
'
'Sub ForeColorOut()
'    Dim ctrl As Control
'    For Each ctrl In mFrame.Controls
'        If InStr(1, ctrl.Name, "DateLabel") <> 0 Then
'            ctrl.ForeColor = vbGrayText
'        End If
'    Next
'End Sub

Private Sub DateLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Unload DatePicker
End Sub

Private Sub DateLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMoveIcon
End Sub

Private Sub MonthLabel_Click()
    
End Sub

Private Sub NextButton_Click()
    With MonthLabel
        If .Tag = 12 Then
            .Tag = 0
            YearLabel = YearLabel + 1
        End If
        
        .Tag = .Tag + 1
        .Caption = MonthName(.Tag)
        CalendarCreate YearLabel, .Tag
    End With
End Sub

Private Sub PreviewButton_Click()
    With MonthLabel
        If .Tag = 1 Then
            .Tag = 13
            YearLabel = YearLabel - 1
        End If
        
        .Tag = .Tag - 1
        .Caption = MonthName(.Tag)
        CalendarCreate YearLabel, .Tag
    End With
End Sub

Private Sub PreviewButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMoveIcon
End Sub

Private Sub NextButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MouseMoveIcon
End Sub

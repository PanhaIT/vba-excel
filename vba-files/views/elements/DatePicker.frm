VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker 
   Caption         =   "DatePicker"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12645
   OleObjectBlob   =   "DatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    setCalendarProps txtFirstWeek, txtDateType, cbCalLanguage
End Sub

Private Sub lblBack_Click()
    FrameSettings.Visible = False
End Sub

Private Sub lblSettingIcon_Click()
    With FrameSettings
        .Left = 6
        .Top = 12
        .Visible = True
        
    End With
    txtFirstWeek = FirstWeek
    txtDateType = DateType
    cbCalLanguage = CalLanguage
End Sub

Private Sub lblToday_Click()
    calendar.CalendarCreate Year(Date), Month(Date)
    calendar.SelectDay (Date)
End Sub

Private Sub lblExit_Click()
    CurrentDate = ""
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Width = 320
        .Height = 340
    End With
 '//Calendar Settings
    LoadLanguagePreference
    
    IconDesign lblSettingIcon, "0071"
    IconDesign lblBack, "0062"
    cbCalLanguage.List = Array("TR", "EN", "PT", "FR")
    
End Sub

Function ChooseDate(dateInput As control)
    Dim cYear As Integer
    Dim cMonth As Integer
    Dim cDay As Date

    With Me
        
         If IsDate(dateInput) Then
             cYear = Year(dateInput)
             cMonth = Month(dateInput)
             cDay = dateInput
         Else
             cYear = Year(Date)
             cMonth = Month(Date)
             cDay = Date
         End If
             calendar.CalendarCreate cYear, cMonth, cDay
             calendar.SelectDay (cDay)
            .Show
         
         If CurrentDate <> Empty Then
             dateInput = CurrentDate
         End If
         
    End With
End Function

Sub IconDesign(ctrl As MSForms.Label, IconCode As String, Optional FontSize As Integer = 14)
    With ctrl
        .Font.Name = "myicons"
        .Font.Size = FontSize
        .ForeColor = RGB(140, 140, 140)
        .Caption = ChrW("&H" & IconCode)
        .BackStyle = fmBackStyleTransparent
        .AutoSize = True
        .WordWrap = False
    End With
End Sub

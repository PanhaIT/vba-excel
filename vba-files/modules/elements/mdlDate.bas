Attribute VB_Name = "mdlDate"
Private Declare PtrSafe Function GetUserDefaultUILanguage Lib "kernel32.dll" () As Long
Public calendar         As New clsDate
Public ActiveDate       As MSForms.Label
Public OldForeColor     As Long

Public FirstWeek        As Integer
Public DateType         As String
Public CalLanguage      As String

Sub setCalendarProps(Optional FirstWeekVal As Integer = 2, _
                     Optional DateTypeVal As String = "dd.mm.yyyy", _
                     Optional language As String = "TR")
                     
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\requires\CalendarSettings.txt"  ' Ayar dosyas�n�n yolu
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, "FirstWeek:" & FirstWeekVal & _
                   ";DateType:" & DateTypeVal & _
                   ";CalLanguage:" & language
    Close #fileNum
    LoadLanguagePreference
End Sub

Sub LoadLanguagePreference()
    Dim filePath As String
    Dim arr
    Dim arrValue
    Dim fileNum As Integer

    filePath = ThisWorkbook.Path & "\requires\CalendarSettings.txt" ' Ayar dosyas�n�n yolu
    fileNum = FreeFile
    
    If Dir(filePath) <> "" Then
        Open filePath For Input As #fileNum
        Line Input #fileNum, arr
        Close #fileNum
        arr = Split(arr, ";")
    End If

    For i = 0 To UBound(arr)
        arrValue = Split(arr(i), ":") 
        If arrValue(0) = "FirstWeek" Then FirstWeek = arrValue(1)
        If arrValue(0) = "DateType" Then DateType = Trim(arrValue(1))
        If arrValue(0) = "CalLanguage" Then CalLanguage = Trim(arrValue(1))

    Next

    With DatePicker
        Select Case CalLanguage
           
            Case Is = "TR"
                .lblToday = "Bug�n"
                .lblExit = "�IK"
            
            Case Is = "EN"
                .lblToday = "Today"
                .lblExit = "Exit"
            Case Is = "PT"
                .lblToday = "Hoje"
                .lblExit = "Sa�da"
        End Select
    End With
End Sub

Function GetSystemLanguage() As String
    Dim langCode As Long
    langCode = GetUserDefaultUILanguage()
    
    ' LangID'in d���k 10 biti dil kodunu, �st 6 biti karakter setini temsil eder
    Dim langID As Integer
    langID = langCode And &H3FF
    
    Select Case langID
        Case &H40 ' �ngilizce
            GetSystemLanguage = "English"
        Case &H41 ' Arap�a
            GetSystemLanguage = "Arabic"
        Case &H42 ' Bulgarca
            GetSystemLanguage = "Bulgarian"
        ' ... Di�er diller burada eklenebilir
        Case Else
            GetSystemLanguage = langID
    End Select
End Function

Sub TestGetSystemLanguage()
    Dim sysLanguage As String
    sysLanguage = GetSystemLanguage()
    
    MsgBox "Sistem Dili: " & sysLanguage
End Sub


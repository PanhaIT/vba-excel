Attribute VB_Name = "mdlPublic"
'-----------------------------------------------------------------------------------------------------------
' #If VBA7 Then
'     Public Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
'     Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
' #Else
'     Public Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
'     Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
' #End If

' Private Const IDC_HAND = 32649&
' Public ListRows
' Public PassiveTabLabel As Object
' Public ActiveTabLabel As Object

' Function ListRowCount(mFrame As MSForms.Frame)
'     ListRowCount = Application.WorksheetFunction.RoundDown(mFrame.Height / 43, 0)
' '    ListRowCount = 15
' End Function

' Public Function MouseCursor(CursorType As Long)
'   Dim lngRet As Long
'   lngRet = LoadCursorBynum(0&, CursorType)
'   lngRet = SetCursor(lngRet)
' End Function

' Public Function MouseMoveIcon()
'     Call MouseCursor(IDC_HAND)
' End Function
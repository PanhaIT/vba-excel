Attribute VB_Name = "mdlDeclareFontUnicode"

Option Explicit

Private Declare PtrSafe Function MessageBoxW Lib "User32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Public Function MsgBoxW(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "Dialog Information") As VbMsgBoxResult
    MsgBoxW = MessageBoxW(Application.hWnd, StrPtr(Prompt), StrPtr(Title), Buttons)
End Function

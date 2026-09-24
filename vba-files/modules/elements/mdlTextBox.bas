Attribute VB_Name = "mdlTextBox"
Public fColor, eColor, tColor, aColor, sColor
Public Const FontSize = 11
Public Const FontName = "Khmer OS Battambang"
Public Const IconFont = "myicons"
Public textControl, ActiveBox, BType
'Public UserFormActive

Public Sub TxtColor(Optional ForeColor As String = "#383838", _
                    Optional EnterColor As String = "#005ea2", _
                    Optional TitleColor As String = "#787875", _
                    Optional AlertColor As String = "#f72111", _
                    Optional SuccessColor As String = "#22ab2b")
    
                    fColor = GetRGBFromHex(ForeColor)
                    eColor = GetRGBFromHex(EnterColor)
                    tColor = GetRGBFromHex(TitleColor)
                    aColor = GetRGBFromHex(AlertColor)
                    sColor = GetRGBFromHex(SuccessColor)
End Sub

Private Function GetRGBFromHex(hexColor As String) As String
        hexColor = VBA.Replace(hexColor, "#", "")
        hexColor = VBA.right("000000" & hexColor, 6)
        
        Red = VBA.Val("&H" & VBA.Mid(hexColor, 1, 2))
        Green = VBA.Val("&H" & VBA.Mid(hexColor, 3, 2))
        blue = VBA.Val("&H" & VBA.Mid(hexColor, 5, 2))

        GetRGBFromHex = RGB(Red, Green, blue)
End Function

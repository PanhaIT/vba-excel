Attribute VB_Name = "add_seller"

Public  Sub saveSeller(sellerId As String)
    Dim sellerName, telephone, address, id, sellerCode, sellerOption, sellerNameEn, email, note As String
    Dim i, LastRow As Long
    Dim ws As Worksheet
    Dim seller_data As Range

    Set ws = Sheets("Seller")

    With frmAddSeller
        sellerName   = .seller_name
        sellerNameEn = .seller_name_en
        telephone    = .telephone
        email        = .email
        address      = .address
        note         = .note
    End With

    If (sellerName = "" Or sellerNameEn = "") Then
        MsgBox "Please input seller name, try again."
        Exit Sub
    Else
        Set seller_data = ws.Range("C7:K7")
        id              = ws.Cells(Rows.Count, "C").End(xlUp).Row - 5
        sellerCode      = "KS" & PadStr(id, 3, "0", xlHAlignRight)
        sellerOption    = PadStr(id, 3, "0", xlHAlignRight) & "." & sellerName

        seller_data.Cells(id, 1)  = id  'id
        seller_data.Cells(id, 2)  = sellerCode  'code
        seller_data.Cells(id, 3)  = sellerName  'seller name
        seller_data.Cells(id, 4)  = sellerNameEn  'seller name latin
        seller_data.Cells(id, 5)  = telephone  'phone
        seller_data.Cells(id, 6)  = email  'social network link
        seller_data.Cells(id, 7)  = address  'address
        seller_data.Cells(id, 8)  = note  'note
        seller_data.Cells(id, 9)  = 1  'status
        seller_data.Cells(id, 10) = sellerOption 'member name option

        Sheets("AddQuotation").Range("quot_seller").Value = sellerOption
        With frmAddSeller
            .seller_name = ""
            .seller_name_en = ""
            .telephone = ""
            .email = ""
            .address = ""
            .note = ""
        End With
        frmAddSeller.Hide
        MsgBoxW "New seller " & sellerName & " (" & sellerNameEn & ") already saved."
        Exit Sub
    End If
End Sub
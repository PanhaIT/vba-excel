Attribute VB_Name = "add_ship_to"

Public  Sub saveShipTo(id As String)
    Dim shipToCode, telephone, address As String
    Dim i, LastRow As Long
    Dim ws As Worksheet
    Dim ship_to_data As Range

    Set ws = Sheets("ShipTo")
    telephone = frmAddShipTo.telephone
    address = frmAddShipTo.address

    If (address = "" Or telephone = "") Then
        MsgBox "Please input require field, try again."
        Exit Sub
    Else
        i = 1
        Set ship_to_data = ws.Range("C6:G6")
        LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row + 1
        Do Until WorksheetFunction.CountA(ship_to_data.Rows(i)) = 0
            i = i + 1 'Last row
        Loop

        shipToCode = PadStr(i, 5, "0", xlHAlignRight) 'PadStr call from declareFunction
        ship_to_data.Cells(i, 1) = i 'id
        ship_to_data.Cells(i, 2) = "ST" & shipToCode 'code
        ship_to_data.Cells(i, 3) = address 'address name
        ship_to_data.Cells(i, 4) = telephone 'address name
        ship_to_data.Cells(i, 5) = 1 'status

        'reset dialog
        With frmAddShipTo
            .telephone = ""
            .address = ""
        End With
        Sheets("AddQuotation").Range("quot_ship_to").Value = "ST" & shipToCode & "." & address

        frmAddShipTo.Hide
        MsgBox "New shipment address already saved."
        Exit Sub
    End If
End Sub
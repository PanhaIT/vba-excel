
Attribute VB_Name = "add_customer"

Private Sub btnAddCustomer()
    With frmCustomer
        .id = 0
        .Show
    End With
End Sub

Sub SaveCustomer(customerId As String)
   ' address,payment_term_id,payment_every,customer_code,name, name_kh, main_number, photo, email, vat , remark, cGroup
    Dim cellCus, sheetName, customerCode, customerNameKh, customerName, address, mainNumber, email, remark, cGroup, paymentTerm, netDay, vat, id, cellId, invoiceSheet, socialNetwork As String
    Dim i, LastRow As Long
    Dim ws As Worksheet
    Dim customer_data As Range
    Dim answer As VbMsgBoxResult

    If (ActiveSheet.name = "CustomerList") Then
        sheetName = "CustomerList"
        cellCus = ""
    ElseIf (ActiveSheet.name = "AddInvoice") Then
        sheetName = "AddInvoice"
        cellCus = "invoice_customer"
    ElseIf (ActiveSheet.name = "AddQuotation") Then
        sheetName = "AddQuotation"
        cellCus = "quot_customer"
    End If

    Set ws = Sheets(sheetName)
    
    With frmCustomer
        customerCode = .customer_code
        customerNameKh = .name_kh
        customerName = .customer_name
        mainNumber = .main_number
        email = .email
        address = .address
        remark = .remark
        cGroup = .cGroup
        paymentTerm = .paymentTerm
        netDay = .payment_every
        vat = .vat
        socialNetwork = .social_network
        ' photo         = .photo
        ' Debug.Print  "customerCode=" & customerCode & ", customerNameKh=" & customerNameKh & ", customerName=" & customerName & ", mainNumber=" & mainNumber & ", email=" & email
        ' Debug.Print  "address=" & address & ", remark=" & remark & ", cGroup=" & cGroup & ", paymentTerm=" & paymentTerm & ", netDay=" & netDay & ", vat=" & vat
    End With
    
    If (customerCode = "" Or customerNameKh = "" Or customerName = "" Or mainNumber = "" Or address = "" Or cGroup = "") Then
        MsgBox "Please input require field, try again."
        Exit Sub
    Else
        i = 1
        Set customer_data = Sheets("CustomerList").Range("C5:Q5")
        LastRow = Sheets("CustomerList").Cells(Rows.Count, "C").End(xlUp).Row + 1
        Do Until WorksheetFunction.CountA(customer_data.Rows(i)) = 0
            i = i + 1 'Last row
        Loop
        id = PadStr(i, 4, "0", xlHAlignRight)
        cellId = "C" & LastRow
        customer_data.Cells(i, 1) = id   'id
        customer_data.Cells(i, 2) = customerCode  'customer code
        customer_data.Cells(i, 3) = customerNameKh  'customer company
        customer_data.Cells(i, 4) = cGroup  'customer group
        customer_data.Cells(i, 5) = customerName  'customer name
        customer_data.Cells(i, 6) = customerNameKh  'customer name kh
        customer_data.Cells(i, 7) = mainNumber  'telephone
        customer_data.Cells(i, 8) = socialNetwork  'social network link
        customer_data.Cells(i, 9) = address  'address
        customer_data.Cells(i, 10) = email 'email
        customer_data.Cells(i, 11) = remark 'remark
        customer_data.Cells(i, 12) = vat 'VAT 
        customer_data.Cells(i, 13) = "=SUMIFS([database_kscm.xlsm]SaleInvoiceData!$AC$8:$AC$10000,[database_kscm.xlsm]SaleInvoiceData!$N$8:$N$10000," & cellId & ",[database_kscm.xlsm]SaleInvoiceData!$AD$8:$AD$10000,"">0"")" 'total balance
        customer_data.Cells(i, 14) = 1 'status
        customer_data.Cells(i, 15) = id & "." & customerNameKh & " (" & customerName & ")" 'customer name option

        reset
    
        If (cellCus <> "") Then
            ws.Range(cellCus).Value = id & "." & customerNameKh & " (" & customerName & ")"
        End If
        
        frmCustomer.Hide
        MsgBox "New customer " & customerCode & "-" & customerName & " already saved."
        Exit Sub
    End If
End Sub

Private Sub reset()
    With frmCustomer
        .customer_code = ""
        .name_kh = ""
        .customer_name = ""
        .main_number = ""
        .email = ""
        .address = ""
        .remark = ""
        .cGroup = ""
        .paymentTerm = ""
        .payment_every = ""
        .limit_balance = ""
        .limit_total_invoice = ""
        .vat = ""
        .social_network = ""
    End With
End Sub

Attribute VB_Name = "add_vendor"

Private Sub btnAddVendor()
    With frmVendor
        .id = 0
        .Show
    End With
End Sub

Public Sub SaveVendor(VendorId As String)
     ' 1. Turn off background processes to maximize speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

        On Error GoTo ErrorHandler ' Ensures settings restore if code crashes

        ' ==========================================
        ' PLACE YOUR DATA WRITING & SAVE CODE HERE
        Dim cellCus, sheetName, VendorCode, VendorNameKh, VendorName, address, mainNumber, email, note, contactPosition, paymentTerm, netDay, tinNumber, id, cellId, invoiceSheet As String
        Dim i, LastRow As Long
        Dim ws As Worksheet
        Dim vendor_data As Range
        Dim answer As VbMsgBoxResult
        Dim otherNumber,vendorContact,faxNumber,socialNetwork As String
        Dim createdDate As Date
        createdDate = Format(Now(), "yyyy/mm/dd hh:mm:ss")

        If (ActiveSheet.name = "VendorList") Then
            sheetName = "VendorList"
            cellCus   = ""
        ElseIf (ActiveSheet.name = "AddPurchaseBill") Then
            sheetName = "AddPurchaseBill"
            cellCus   = "pb_vendor"
        End If

        Set ws = Sheets(sheetName)
        
        With frmVendor
            VendorCode      = .vendor_code
            VendorNameKh    = .vendor_name_kh
            VendorName      = .vendor_name
            vendorContact   = .vendor_contact
            mainNumber      = .main_number
            otherNumber     = .other_telephone
            faxNumber       = .fax_number
            email           = .email
            address         = .address
            note            = .note
            contactPosition = .contactPosition
            paymentTerm     = .paymentTerm
            netDay          = .payment_every
            tinNumber       = .tin_number
            socialNetwork   = .social_network
            'photo          = .photo
        End With
        
        If (VendorCode = "" Or VendorNameKh = "" Or VendorName = "" Or mainNumber = "") Then
            MsgBox "Please input require field, try again."
            Exit Sub
        Else
            i = 1
            Set vendor_data = Sheets("VendorList").Range("C8:T8")
            LastRow         = Sheets("VendorList").Cells(Rows.Count, "C").End(xlUp).Row + 1
            Do Until WorksheetFunction.CountA(vendor_data.Rows(i)) = 0
                i = i + 1 'Last row
            Loop

            id     = PadStr(i, 3, "0", xlHAlignRight)
            cellId = "C" & LastRow

            vendor_data.Cells(i, 1)  = id  'id
            vendor_data.Cells(i, 2)  = ""  'Photo
            vendor_data.Cells(i, 3)  = VendorCode  'Vendor code
            vendor_data.Cells(i, 4)  = VendorNameKh  'Vendor company
            vendor_data.Cells(i, 5)  = VendorName  'Vendor Name En
            vendor_data.Cells(i, 6)  = vendorContact  'Vendor Contact
            vendor_data.Cells(i, 7)  = contactPosition 'Contact Position
            vendor_data.Cells(i, 8)  = mainNumber  'Work Telephone
            vendor_data.Cells(i, 9)  = otherNumber  'Other Telephone
            vendor_data.Cells(i, 10) = faxNumber  'Fax Number
            vendor_data.Cells(i, 11) = email 'email
            vendor_data.Cells(i, 12) = tinNumber 'tinNumber
            vendor_data.Cells(i, 13) = socialNetwork 'Social Network
            vendor_data.Cells(i, 14) = address  'address
            vendor_data.Cells(i, 15) = note 'note
            vendor_data.Cells(i, 16) = createdDate 'Created
            vendor_data.Cells(i, 17) = "=IF(" & cellId & "="""","""",SUMIFS([database_kscm.xlsm]PurchaseBill!$AE$8:$AE$10000,[database_kscm.xlsm]PurchaseBill!$P$8:$P$10000," & cellId & ",[database_kscm.xlsm]PurchaseBill!$AM$8:$AM$10000,"">0""))"
            ' vendor_data.Cells(i, 17) = "=SUMIFS([database_kscm.xlsm]PurchaseBill!$AE$8:$AE$10000,[database_kscm.xlsm]PurchaseBill!$P$8:$P$10000," & cellId & ",[database_kscm.xlsm]PurchaseBill!$AM$8:$AM$10000,"">0"")" 'Vendor Balance
            vendor_data.Cells(i, 18) = 1 'Is Active
            vendor_data.Cells(i, 19) = id & "." & VendorNameKh & " (" & VendorName & ")" 'Vendor name option

            reset

            If (cellCus <> "") Then
                ws.Range(cellCus).Value = id & "." & VendorNameKh & " (" & VendorName & ")"
            End If

            frmVendor.Hide
            MsgBox "New Vendor " & VendorCode & "-" & VendorName & " already saved."
            Exit Sub
        End If
        ' ==========================================

    CleanUp:
        ' 2. Turn settings back on after saving is finished
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Exit Sub

    ErrorHandler:
        MsgBox "An error occurred: " & Err.Description
        Resume CleanUp
    
End Sub

Private Sub reset()
    With frmVendor
        .vendor_code = ""
        .vendor_name = ""
        .vendor_name_kh =""
        .vendor_contact = ""
        .contactPosition =""
        .main_number = ""
        .other_telephone = ""
        .fax_number = ""
        .email = ""
        .address = ""
        .note = ""
        .tin_number = ""
        .paymentTerm = ""
        .payment_every = ""
        .social_network = ""
    End With
End Sub
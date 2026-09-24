Attribute VB_Name = "lock_protected_sheet_cell"

Sub LockAddInvoiceSheet()
    Dim activeWorkbook As Workbook
    Dim ws_add_inv As Worksheet
    Dim messageAlertBox AS VbMsgBoxResult
    messageAlertBox = MsgBox("Are you want to lock and protect sheet AddInvoice?", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Lock Formula Add Invoice")

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_add_inv      = activeWorkbook.Sheets("AddInvoice")

    If (messageAlertBox = vbYes) Then 
        ' 1. Unprotect sheet to make structural changes
        ws_add_inv.Unprotect Password:="kscm@2026"
        
        ' 2. Unlock EVERY cell first
        With ws_add_inv
            .Cells.Locked = False
        End With
        
        ' 3. Lock only your specific range
        With ws_add_inv
            .Range("invoice_code").Locked= True
            .Range("invoice_no").Locked = True
            .Range("invoice_customer_id").Locked= True
            .Range("invoice_branch_id").Locked = True
            .Range("invoice_location_group_id").Locked= True
            .Range("invoice_location_id").Locked = True
            .Range("invoice_price_type_id").Locked= True
            .Range("invoice_discount_id").Locked = True
            .Range("invoice_seller_id").Locked= True
            .Range("invoice_type_id").Locked = True
            .Range("invoice_payment_term_id").Locked= True
            .Range("invoice_currency_id").Locked = True
            .Range("invoice_week").Locked= True
            .Range("invoice_vat_percent").Locked = True
            .Range("invoice_vat_cal_id").Locked = True
            .Range("invoice_total_vat").Locked= True
            .Range("invoice_deposit").Locked = True
            .Range("invoice_discount").Locked= True
            .Range("invoice_sub_total_amount").Locked = True
            .Range("invoice_status").Locked= True
            .Range("invoice_total_amount").Locked = True

            .Range("I8:L87").Locked = True
            .Range("S8:S87").Locked = True
            .Range("U8:V87").Locked = True
            .Range("X8:Y87").Locked = True
            .Range("AA8:BB87").Locked = True
        End With
        
        ' 4. Protect the sheet to enforce the lock
        ws_add_inv.Protect Password:="kscm@2026"
    Else
        Exit Sub
    End If
End Sub

Sub unlockAddInvoiceSheet()
    frmLockProtectSheet.Show
End Sub
Attribute VB_Name = "lock_protect_add_quote"

Public  Sub LockAddQuoteSheet()
    Dim activeWorkbook As Workbook
    Dim ws_add_inv As Worksheet
    Dim messageAlertBox AS VbMsgBoxResult
    messageAlertBox = MsgBox("Are you want to lock and protect sheet AddQuotation?", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Lock Formula Add Quote")

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_add_inv      = activeWorkbook.Sheets("AddQuotation")

    If (messageAlertBox = vbYes) Then 
        ' 1. Unprotect sheet to make structural changes
        ws_add_inv.Unprotect Password:="kscm@2026"
        
        ' 2. Unlock EVERY cell first
        With ws_add_inv
            .Cells.Locked = False
        End With
        
        ' 3. Lock only your specific range
        With ws_add_inv
            .Range("quot_code").Locked= True
            .Range("quot_no").Locked = True
            .Range("quot_customer_id").Locked= True
            .Range("quot_branch_id").Locked = True
            .Range("quot_location_group_id").Locked= True
            .Range("quot_location_id").Locked = True
            .Range("quot_price_type_id").Locked= True
            .Range("quot_discount_id").Locked = True
            .Range("quot_seller_id").Locked= True
            .Range("quot_currency_id").Locked = True
            .Range("quot_week").Locked= True
            .Range("quot_vat_percent").Locked = True
            .Range("quot_vat_cal_id").Locked = True
            .Range("quot_total_vat").Locked= True
            .Range("quot_deposit").Locked = True
            .Range("quot_discount").Locked= True
            .Range("quot_sub_total_amount").Locked = True
            .Range("quot_status").Locked= True
            .Range("quot_total_amount").Locked = True
            .Range("quot_currency_symbol").Locked = True
            .Range("quot_ship_telephone").Locked = True
            .Range("F11").Locked = True
            .Range("AA90").Locked = True
            
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

Public  Sub unlockAddQuoteSheet()
    frmLockProtectSheet.Show
End Sub

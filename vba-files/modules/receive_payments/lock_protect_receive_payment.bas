Attribute VB_Name = "lock_protect_receive_payment" 

Public  Sub LockProtectFormulaReceivePayment()
    Dim activeWorkbook As Workbook
    Dim ws_add_inv As Worksheet
    Dim messageAlertBox AS VbMsgBoxResult
    messageAlertBox = MsgBox("Are you want to lock and protect sheet AddReceivePayment?", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Lock Formula Add Receive Payment")

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_add_inv      = activeWorkbook.Sheets("AddReceivePayment")

    If (messageAlertBox = vbYes) Then 
        ' 1. Unprotect sheet to make structural changes
        ws_add_inv.Unprotect Password:="kscm@2026"
        
        ' 2. Unlock EVERY cell first
        With ws_add_inv
            .Cells.Locked = False
        End With

        ' 3. Lock only your specific range
        With ws_add_inv
            .Range("rp_invoice_id").Locked = True
            .Range("rp_invoice_code").Locked= True
            .Range("rp_invoice_date").Locked = True
            .Range("rp_customer_name").Locked= True
            .Range("rp_chart_account").Locked= True
            .Range("rp_chart_account_id").Locked = True
            .Range("rp_customer_id").Locked= True
            .Range("rp_currency_id").Locked = True
            .Range("rp_payment_net_days").Locked = True
            .Range("rp_week").Locked= True
            .Range("U4:W4").Locked = True

            '************Print_Area************
            ' .Range("rp_receipt_code").Locked= True
            ' .Range("rp_amount_due").Locked = True
            ' .Range("rp_total_paid").Locked= True
            ' .Range("rp_amount_paid_before").Locked = True
            ' .Range("rp_total_balance").Locked= True
            ' .Range("rp_total_amount_paid").Locked= True
            .Range("M22:X29").Locked = True
            .Range("Print_Area").Locked = True
            .Range("Z2:AH5000").Locked = True
        End With
        
        ' 4. Protect the sheet to enforce the lock
        ws_add_inv.Protect Password:="kscm@2026"
    Else
        Exit Sub
    End If
End Sub

Public  Sub unlockFormulaReceivePayment()
    frmLockProtectSheet.Show
End Sub

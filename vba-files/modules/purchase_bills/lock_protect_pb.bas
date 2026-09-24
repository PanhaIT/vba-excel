Attribute VB_Name = "lock_protect_pb" 

Sub LockAddPurchaseBillSheet()
    Dim activeWorkbook As Workbook
    Dim ws_add_inv As Worksheet
    Dim messageAlertBox AS VbMsgBoxResult
    messageAlertBox = MsgBox("Are you want to lock and protect sheet AddPurchaseBill?", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Lock Formula Add Purchase Bill")

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_add_inv      = activeWorkbook.Sheets("AddPurchaseBill")

    If (messageAlertBox = vbYes) Then 
        ' 1. Unprotect sheet to make structural changes
        ws_add_inv.Unprotect Password:="kscm@2026"
        
        ' 2. Unlock EVERY cell first
        With ws_add_inv
            .Cells.Locked = False
        End With
        
        ' 3. Lock only your specific range
        With ws_add_inv
            .Range("pb_no").Locked= True
            .Range("pb_code").Locked = True
            .Range("pb_vendor_id").Locked= True
            .Range("pb_branch_id").Locked = True
            .Range("pb_location_id").Locked= True
            .Range("pb_discount_id").Locked = True
            .Range("pb_payment_term_id").Locked= True
            .Range("pb_currency_id").Locked = True
            .Range("pb_discount_percent").Locked= True
            .Range("pb_week").Locked = True
            .Range("pb_net_days").Locked= True
            .Range("pb_vat_percent").Locked = True
            .Range("pb_vat_cal_id").Locked = True
            .Range("pb_total_vat").Locked = True
            .Range("pb_deposit").Locked= True
            .Range("pb_discount").Locked = True
            .Range("pb_sub_total_amount").Locked= True
            .Range("pb_status").Locked = True
            .Range("pb_total_amount").Locked= True
            
            .Range("I8:L87").Locked = True
            .Range("S8:S87").Locked = True
            .Range("U8:V87").Locked = True
            .Range("X8:Y87").Locked = True
            .Range("AA8:AT87").Locked = True
        End With
        
        ' 4. Protect the sheet to enforce the lock
        ws_add_inv.Protect Password:="kscm@2026"
    Else
        Exit Sub
    End If
End Sub

Sub unlockAddPurchaseBillSheet()
    frmLockProtectSheet.Show
End Sub

' Sub ListAllSheetNames()
'     Dim ws As Worksheet
'     Dim rowNum As Integer
'     rowNum = 1

'     For Each ws In ActiveWorkbook.Worksheets
'         If (ws.Name = "Sheet1") Then 'AddPurchaseBill
'             debug.print ws.Name
'             Debug.Print "*******************" 
'             ws.Protect Password:="kscm@2026"
'         End If
'         ' Cells(rowNum, 1).Value = ws.Name
'         rowNum = rowNum + 1
'     Next ws
' End Sub
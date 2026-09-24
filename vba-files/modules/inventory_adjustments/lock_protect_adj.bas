Attribute VB_Name = "lock_protect_adj"

Sub LockAddAdjustmentSheet()
    Dim activeWorkbook As Workbook
    Dim ws_add_inv As Worksheet
    Dim messageAlertBox AS VbMsgBoxResult
    messageAlertBox = MsgBox("Are you want to lock and protect sheet InventoryAdj?", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Lock Formula Add Adjustment")

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_add_inv      = activeWorkbook.Sheets("InventoryAdj")

    If (messageAlertBox = vbYes) Then 
        ' 1. Unprotect sheet to make structural changes
        ws_add_inv.Unprotect Password:="kscm@2026"
        
        ' 2. Unlock EVERY cell first
        With ws_add_inv
            .Cells.Locked = False
        End With
        
        ' 3. Lock only your specific range
        With ws_add_inv
            .Range("adj_code").Locked = True
            .Range("adj_location_group_id").Locked = True
            .Range("adj_branch_id").Locked = True
            .Range("adj_type").Locked = True
            .Range("adj_check_ref_invoice").Locked = True
            .Range("adj_week").Locked = True
            .Range("adj_chart_account_id").Locked = True
            .Range("adj_status").Locked = True
            .Range("adj_total_qty").Locked = True
            .Range("adj_no").Locked = True
            .Range("last_adj_id").Locked = True
            .Range("adj_created_by").Locked = True
            
            .Range("I8:K87").Locked = True
            .Range("P8:Q87").Locked = True
            .Range("V8:AO87").Locked = True
        End With
        
        ' 4. Protect the sheet to enforce the lock
        ws_add_inv.Protect Password:="kscm@2026"
    Else
        Exit Sub
    End If
End Sub

Sub unlockAddAdjustmentSheet()
    frmLockProtectSheet.Show
End Sub
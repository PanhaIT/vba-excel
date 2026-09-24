VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLockProtectSheet 
   Caption         =   "Unlock And Unprotect Sheet"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9495.001
   OleObjectBlob   =   "frmLockProtectSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLockProtectSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub reset_Click()
    With frmLockProtectSheet
        .lock_password              = ""
        .add_inv_adj.Value          = False
        .add_purchase_bill.Value    = False
        .add_sales_invoice.Value    = False
        .add_quotation.Value        = False
    End With
End Sub

Private Sub Cancel_Click()
    With frmLockProtectSheet
        .lock_password              = ""
        .add_inv_adj.Value          = False
        .add_purchase_bill.Value    = False
        .add_sales_invoice.Value    = False
        .add_quotation.Value        = False
    End With
    Unload Me
End Sub

Private Sub save_Click()
    Dim activeWorkbook As Workbook
    Dim ws_add_adj, ws_add_inv, ws_add_pb, ws_add_quote, ws_add_rp As Worksheet

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_add_adj      = activeWorkbook.Sheets("InventoryAdj")
    Set ws_add_inv      = activeWorkbook.Sheets("AddInvoice")
    Set ws_add_pb       = activeWorkbook.Sheets("AddPurchaseBill")
    Set ws_add_quote    = activeWorkbook.Sheets("AddQuotation")
    Set ws_add_rp       = activeWorkbook.Sheets("AddReceivePayment")

    ' 1. Unprotect sheet to make structural changes
    If (Me.lock_password = "kscm@2026") Then

        If Me.add_inv_adj.Value = True Then
            ' Debug.Print "The add_inv_adj checkbox has been CHECKED."
            ws_add_adj.Unprotect Password:=Me.lock_password  ' password = "kscm@2026"
            With ws_add_adj
                .Cells.Locked = False
            End With
        End If

        If Me.add_purchase_bill.Value = True Then
            ' Debug.Print "The add_purchase_bill checkbox has been CHECKED."
            ws_add_pb.Unprotect Password:=Me.lock_password  ' password = "kscm@2026"
            With ws_add_pb
                .Cells.Locked = False
            End With
        End If

        If Me.add_sales_invoice.Value = True Then
            ' Debug.Print "The add_sales_invoice checkbox has been CHECKED."
            ws_add_inv.Unprotect Password:=Me.lock_password  ' password = "kscm@2026"
            With ws_add_inv
                .Cells.Locked = False
            End With
        End If

        If Me.add_quotation.Value = True Then
            ' Debug.Print "The add_sales_invoice checkbox has been CHECKED."
            ws_add_quote.Unprotect Password:=Me.lock_password  ' password = "kscm@2026"
            With ws_add_quote
                .Cells.Locked = False
            End With
        End If
        
        If Me.add_receive_payment.Value = True Then
            ' Debug.Print "The add_sales_invoice checkbox has been CHECKED."
            ws_add_rp.Unprotect Password:=Me.lock_password  ' password = "kscm@2026"
            With ws_add_rp
                .Cells.Locked = False
            End With
        End If

        Unload Me
    Else
        MsgBox "Incorrect password , please try to input again."
        Exit Sub
    End If
End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Initialize()
    TxtColor ForeColor:="#43545F", _
        EnterColor:="#005ea2", _
        TitleColor:="#ababab", _
        AlertColor:="#f72111", _
        SuccessColor:="#22ab2b"
    tbox.clasBox Me, "Border"
    cBtn.classButton Me
End Sub


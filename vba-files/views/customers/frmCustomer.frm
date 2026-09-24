VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustomer 
   Caption         =   "Add Customer"
   ClientHeight    =   8910.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21045
   OleObjectBlob   =   "frmCustomer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub save_Click()
    tbox.TxtControl
    If textControl = True Then
        SaveCustomer id
    End If
End Sub

Private Sub reset_Click()
   resetField
End Sub

Private Sub Cancel_Click()
    resetField
    Unload Me
End Sub

Private Sub resetField()
    With frmCustomer
        ' .customer_code = ""
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

Private Sub UserForm_Activate()
    'Generate auto customer code
    Dim i As Integer
    Dim customerList As Range

    Set ws_cus_list = Sheets("CustomerList")
    Set customerList = ws_cus_list.Range("customer_list")
    LastRow = ws_cus_list.Cells(Rows.Count, "C").End(xlUp).Row + 1

    i = 1
    Do Until WorksheetFunction.CountA(customerList.Rows(i)) = 0
        i = i + 1 'Last row
    Loop

    With frmCustomer
        .customer_code = "CKS" & PadStr(i, 7, "0", xlHAlignRight)
        .id = 0
    End With
End Sub

Private Sub UserForm_Initialize()
    TxtColor ForeColor:="#43545F", _
        EnterColor:="#005ea2", _
        TitleColor:="#ababab", _
        AlertColor:="#f72111", _
        SuccessColor:="#22ab2b"
    tbox.clasBox Me, "Border"
    cBtn.classButton Me
    ' paymentTerm.List = Array("COD", "3 Month") 'paymentTerm is text box name , use to pass option to combo box dropdown list

    Dim ws, ws_setting, ws_cus_list As Worksheet
    Dim rangeCgroup, rowCgroup, paymentTermList, paymentTermOption As Range
    Set ws = ThisWorkbook.Sheets("Cgroup")
    Set ws_setting = ThisWorkbook.Sheets("Setting")
    Dim i, j As Integer
    
    Me.cGroup.Clear
    Set rangeCgroup = ws.Range("cgroup_list")
    Set rowCgroup = ws.Range("cgroup_option")
    For i = 1 To rangeCgroup.Rows.Count
        If rowCgroup.Cells(i) <> "" Then
            Me.cGroup.AddItem rowCgroup(i)
        End If
    Next i

    Me.paymentTerm.Clear
    Set paymentTermList = ws_setting.Range("payment_term_list")
    Set paymentTermOption = ws_setting.Range("payment_term_option")
    For j = 1 To paymentTermList.Rows.Count
        If paymentTermOption.Cells(j) <> "" Then
            Me.paymentTerm.AddItem paymentTermOption(j)
        End If
    Next j
End Sub

Private Sub paymentTerm_Change()
    Dim ws_setting As Worksheet
    Dim paymentTermList As Range
    Dim i As Integer
    Dim paymentId As String
    
    Set ws_setting = ThisWorkbook.Sheets("Setting")
    Set paymentTermList = ws_setting.Range("payment_term_list")
    
    If Me.paymentTerm.Value <> "" Then
        paymentId = Split(Me.paymentTerm.Value, ".")(0) * 1
       For i = 1 To paymentTermList.Rows.Count
            If paymentTermList.Cells(i, 1) = paymentId Then
                Me.payment_every.Value = paymentTermList.Cells(i, 3)
            End If
        Next i
    End If
End Sub

' Private Sub Worksheet_Activate()
'     Dim ws, ws_inv_data As Worksheet
'     Dim rangeInvoiceList, rowInvoiceCode As Range
'     Set ws = Sheets("PrintInvoice")
'     Set ws_inv_data = Sheets("SaleInvoiceData")
'     Dim i, j As Integer
    
'     ws.InvoiceComboBox.Clear
'     Set rangeInvoiceList = ws_inv_data.Range("invoice_list2")
'     Set rowInvoiceCode = ws_inv_data.Range("invoice_code_option")
'     For i = 1 To rangeInvoiceList.Rows.Count
'         If rowInvoiceCode.Cells(i) <> "" Then
'             ws.InvoiceComboBox.AddItem rowInvoiceCode(i)
'         End If
'     Next i
' End Sub


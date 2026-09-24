VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVendor 
   Caption         =   "Add Vendor"
   ClientHeight    =   8895.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21045
   OleObjectBlob   =   "frmVendor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub save_Click()
    tbox.TxtControl
    If textControl = True Then
        SaveVendor id
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
    With frmVendor
        .vendor_name = ""
        .vendor_name_kh = ""
        .vendor_contact = ""
        .contactPosition = ""
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

Private Sub UserForm_Activate()
    'Generate auto vendor code
    Dim ws_vendor_list As Worksheet
    Dim i, LastRow As Integer
    Dim RgVendorList As Range

    Set ws_vendor_list = Sheets("VendorList")
    Set RgVendorList = ws_vendor_list.Range("vendor_list")
    LastRow = ws_vendor_list.Cells(Rows.Count, "C").End(xlUp).Row + 1

    i = 1
    Do Until WorksheetFunction.CountA(RgVendorList.Rows(i)) = 0
        i = i + 1 'Last row
    Loop

    With frmVendor
        .vendor_code = "VKS" & PadStr(i, 3, "0", xlHAlignRight)
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

    Dim ws_setting, ws_vendor_list As Worksheet
    Dim positionList, positionOption, paymentTermList, paymentTermOption As Range
    Set ws_setting = ThisWorkbook.Sheets("Setting")
    Dim i, j As Integer
    
    Me.contactPosition.Clear
    Set positionList = ws_setting.Range("position_list")
    Set positionOption = ws_setting.Range("position_option")
    For i = 1 To positionList.Rows.Count
        If positionOption.Cells(i) <> "" Then
            Me.contactPosition.AddItem positionOption(i)
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


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddSeller 
   Caption         =   "Add New Seller"
   ClientHeight    =   9210.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14685
   OleObjectBlob   =   "frmAddSeller.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAddSeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub iconCancel_Click()
    resetField
    Unload Me
End Sub

Private Sub iconReset_Click()
   resetField
End Sub

Private Sub iconSave_Click()
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
    With frmAddSeller
        .seller_name = ""
        .seller_name_en = ""
        .telephone = ""
        .email = ""
        .address = ""
        .note = ""
    End With
End Sub

Private Sub save_Click()
    tbox.TxtControl
    If textControl = True Then
        saveSeller id 'call sub in sellers/add_seller.bas
    End If
End Sub

Private Sub UserForm_Activate()
    frmAddSeller.id = 0
End Sub

Private Sub UserForm_Initialize()
    TxtColor forecolor:="#43545F", _
        EnterColor:="#005ea2", _
        TitleColor:="#ababab", _
        AlertColor:="#f72111", _
        SuccessColor:="#22ab2b"
    tbox.clasBox Me, "Border"
    cBtn.classButton Me
End Sub


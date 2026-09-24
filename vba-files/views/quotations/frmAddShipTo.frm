VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddShipTo 
   Caption         =   "Add New Ship To"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13005
   OleObjectBlob   =   "frmAddShipTo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddShipTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub reset_Click()
   resetField
End Sub

Private Sub Cancel_Click()
    resetField
    Unload Me
End Sub

Private Sub resetField()
    With frmAddShipTo
        .telephone = ""
        .address = ""
    End With
End Sub

Private Sub save_Click()
    tbox.TxtControl
    If textControl = True Then
        saveShipTo id 'call sub in quotations/add_ship_to.bas
    End If
End Sub

Private Sub UserForm_Activate()
    frmAddShipTo.id = 0
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


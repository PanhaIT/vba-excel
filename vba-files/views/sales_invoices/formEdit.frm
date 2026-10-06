VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEdit 
   Caption         =   "Edit movie details"
   ClientHeight    =   6735
   ClientLeft      =   75
   ClientTop       =   300
   ClientWidth     =   6075
   OleObjectBlob   =   "formEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM

Private m_editRow As Long

Public Property Let editRow(newEditRow As Long)
    m_editRow = newEditRow
End Property

' USERFORM EVENTS
Public Sub Fill(Data As Variant)
    textboxTitle.Value = Data(1, 1)
    textboxGenre.Value = Data(1, 2)
    textboxYear.Value = Data(1, 3)
    textboxDirector.Value = Data(1, 4)
    textboxRevenue.Value = Data(1, 5)
    textboxIMDB.Value = Data(1, 6)
    textboxBudget.Value = Data(1, 7)

End Sub

Private Sub buttonClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

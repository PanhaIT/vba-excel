VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormClass 
   Caption         =   "Movie Database"
   ClientHeight    =   12615
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   22410
   OleObjectBlob   =   "UserFormClass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM

' Declare the ListBox variable at the top of the UserForm code
Private WithEvents m_modernListbox As clsModernListbox
Private WithEvents m_brandModernListbox As clsModernListbox
Private m_displayData As Boolean
Private m_brandDisplayData As Boolean

Private Sub save_Click()

End Sub

Private Sub reset_Click()
   resetField
End Sub

Private Sub Cancel_Click()
    Hide
End Sub

Private Sub resetField()
    total_amount.Value = ""
    invoice_date.Value = ""
    invoice_price_type.Value = ""
    branch.Value = ""
    warehouse.Value = ""
    location.Value = ""
End Sub

' LISTBOX EVENT
' Event when item select in the clsModernListbox
Private Sub m_modernListBox_ItemSelected(dataRow As Long)
    m_displayData = True
    If m_displayData = False Then Exit Sub
    
    Dim form As New formEdit
    Dim arr As Variant

    arr = m_modernListbox.GetRow(dataRow)
    Dim pgroupId As integer
    pgroupId = CInt(MID(arrayRowToString(arr, 1),1,3))
    
    ' Call form.Fill(arr)
    ' form.show
End Sub

Private Sub m_brandModernListbox_ItemSelected(dataRow As Long)
    m_brandDisplayData = True
    If m_brandDisplayData = False Then Exit Sub
    
    Dim form As New formEdit
    Dim arr As Variant
    arr = m_brandModernListbox.GetRow(dataRow)
    Dim brandId As integer
    brandId = CInt(MID(arrayRowToString(arr, 1),1,3))
    MsgBox "dataRow=" & dataRow & "=>arrayRowToString= " & arrayRowToString(arr, 1) & ",brandId = " & brandId
    ' Call form.Fill(arr)
    ' form.show
End Sub

' USERFORM EVENTS
Private Sub UserForm_Initialize()
    ' TxtColor forecolor:="#43545F", _
    ' EnterColor:="#005ea2", _
    ' TitleColor:="#ababab", _
    ' AlertColor:="#f72111", _
    ' SuccessColor:="#22ab2b"
    ' tbox.clasBox Me, "Border"
    ' cBtn.classButton Me

    ' OptionMulti.Value = True
    ' CheckBoxHover.Value = True
    ' checkboxScrollbars.Value = False
    ' checkboxSelectOn.Value = True
    textboxRecords.Value = 10
    brandTextboxRecords.Value = 10
    ' checkboxAutoHeight.Value = True
    ' CheckBoxAutoWidth.Value = True

    Set m_modernListbox = Instantiate_clsModernListbox
    Set m_modernListbox.parentFrame = FrameListBox

    With m_modernListbox
        .HoverOn = True 'CheckBoxHover.Value
        .multiSelect = fmMultiSelectExtended 'IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
        ' .columnWidths = "150;150;90;100;90;90;100"
        .columnWidths = "300"
        .ScrollBars = fmScrollBarsNone 'IIf(checkboxScrollbars.Value = True, fmScrollBarsBoth, fmScrollBarsNone)
        .recordsPerPage = textboxRecords.Value
        ' .AutomaticHeight = checkboxAutoHeight.Value
        ' .AutomaticWidth = CheckBoxAutoWidth.Value
        ' .HeaderFieldsFromString = Sheets("Movies").Range("A1").Value & ";" & Sheets("Movies").Range("B1").Value & ";" & Sheets("Movies").Range("C1").Value & ";" & Sheets("Movies").Range("D1").Value & ";" & Sheets("Movies").Range("E1").Value & ";" & Sheets("Movies").Range("F1").Value & ";" & Sheets("Movies").Range("G1").Value & ";" & Sheets("Movies").Range("H1").Value
         .HeaderFieldsFromString = Sheets("Movies").Range("J1").Value
        Dim rg As Range: Set rg = Sheets("Movies").Range("J2:J31")
        .List = rg.Value
    End With
        
    Set m_brandModernListbox = Instantiate_clsModernListbox
    Set m_brandModernListbox.parentFrame = brandFrameListBox

    With m_brandModernListbox
        .HoverOn = True 'CheckBoxHover.Value
        .multiSelect = fmMultiSelectExtended 'IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
        ' .columnWidths = "150;150;90;100;90;90;100"
        .columnWidths = "300"
        .ScrollBars = fmScrollBarsNone 'IIf(checkboxScrollbars.Value = True, fmScrollBarsBoth, fmScrollBarsNone)
        .recordsPerPage = brandTextboxRecords.Value
        ' .AutomaticHeight = checkboxAutoHeight.Value
        ' .AutomaticWidth = CheckBoxAutoWidth.Value
        ' .HeaderFieldsFromString = Sheets("Movies").Range("A1").Value & ";" & Sheets("Movies").Range("B1").Value & ";" & Sheets("Movies").Range("C1").Value & ";" & Sheets("Movies").Range("D1").Value & ";" & Sheets("Movies").Range("E1").Value & ";" & Sheets("Movies").Range("F1").Value & ";" & Sheets("Movies").Range("G1").Value & ";" & Sheets("Movies").Range("H1").Value
         .HeaderFieldsFromString = Sheets("Movies").Range("K1").Value
        Dim rgBrand As Range: Set rgBrand = Sheets("Movies").Range("K2:K31")
        .List = rgBrand.Value
    End With

    ' Show controls
    Call ShowControlSection
    
    ' Set Userform to screen size
    Me.left = 0
    Me.Height = Application.Height
    Me.top = 0
    Me.width = Application.width
 
End Sub

' HELPER
Private Sub ShowControlSection(Optional turnOn As Boolean = True)
    Dim c As control
    For Each c In Controls
        If TypeName(c) = "Frame" And c.tag = "temp" Then
            c.Visible = turnOn
        End If
    Next c
End Sub

' USERFORM CONTROLS EVENTS

' Display all the selected record in a message box
Private Sub buttonSelectedData_Click()
    Dim Data As Variant
    Data = m_modernListbox.selectedItems()
    MsgBox "Selected records: " & Data
    If IsEmpty(Data) Then
        MsgBox "No records selected"
    Else
        MsgBox arrayToString(Data)
    End If
End Sub

' ' Display the first select record in a message box
Private Sub buttonSelectedOne_Click()
    Dim Data As Variant
    Data = m_modernListbox.SelectedItem()
    If IsEmpty(Data) Then
        MsgBox "No record selected"
    Else
            MsgBox arrayRowToString(Data, 1)
    End If

End Sub

' Private Sub buttonUpdateHeight_Click()
'     If Trim(TextBoxHeight.Value) = "" Then Exit Sub
'     m_modernListbox.Height = Trim(TextBoxHeight.Value)
' End Sub

' Private Sub buttonUpdateWidth_Click()
'     If Trim(textboxWidth.Value) = "" Then Exit Sub
'     m_modernListbox.Width = Trim(textboxWidth.Value)
' End Sub

' Private Sub buttonColumnWidths_Click()
'     If Trim(TextBoxWidths.Value) = "" Then Exit Sub
'     m_modernListbox.columnWidths = Trim(TextBoxWidths.Value)
' End Sub

' Private Sub buttonAddHeaders_Click()
'     If Trim(textboxHeaders.Value) = "" Then Exit Sub
'     m_modernListbox.HeaderFieldsFromString = Trim(textboxHeaders.Value)
' End Sub

' Private Sub checkboxAutoHeight_Click()
'     m_modernListbox.AutomaticHeight = checkboxAutoHeight.Value
' End Sub

' Private Sub CheckBoxAutoWidth_Click()
'     m_modernListbox.AutomaticWidth = CheckBoxAutoWidth.Value
' End Sub

' When turned on the ListBox will display an UserForm with the record details when you click on a record.
' It fired the m_modernListBox_ItemSelected() above when a record is clicked
' Private Sub checkboxDisplaySelect_Click()
'     m_displayData = checkboxDisplaySelect.Value
' End Sub

' Selects the record number specificed in the textbox.
' Note this is the row number of the data and not the row number on the screen
' Private Sub textboxSelectItem_Change()
'     If Len(Trim(textboxSelectItem.Value)) = 0 Then Exit Sub
'     MsgBox "Selecting record number: " & Trim(textboxSelectItem.Value)
'     ' Call m_modernListbox.SetSelected(textboxSelectItem.Value, checkboxSelectOn.Value)
' End Sub

' Turn on/off hover functionality
' Private Sub CheckBoxHover_Click()
'     m_modernListbox.HoverOn = CheckBoxHover.Value
' End Sub

' ' Turn on/off scrollbars
' Private Sub checkboxScrollbars_Click()
'     m_modernListbox.ScrollBars = IIf(checkboxScrollbars.Value = True, fmScrollBarsBoth, fmScrollBarsNone)
' End Sub

' set the number of records displayed on a page
Private Sub textboxRecords_Change()
    ' If Len(Trim(textboxRecords.Value)) = 0 Then Exit Sub
    ' m_modernListbox.recordsPerPage = Trim(textboxRecords.Value)
End Sub

Private Sub brandTextboxRecords_Change()
    ' If Len(Trim(brandTextboxRecords.Value)) = 0 Then Exit Sub
    ' m_brandModernListbox.recordsPerPage = Trim(brandTextboxRecords.Value)
End Sub

Private Sub Frame1_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If Not m_modernListbox Is Nothing Then
        Call m_modernListbox.ClearHover
    End If
End Sub

Private Sub Frame2_Exit(ByVal cancel As MSForms.ReturnBoolean)
    If Not m_brandModernListbox Is Nothing Then
        Call m_brandModernListbox.ClearHover
    End If
End Sub

Private Sub Frame1_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not m_modernListbox Is Nothing Then
        Call m_modernListbox.ClearHover
    End If
End Sub

Private Sub Frame2_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not m_brandModernListbox Is Nothing Then
        Call m_brandModernListbox.ClearHover
    End If
End Sub

' Allow Multiple selections
' Private Sub OptionMulti_Click()
'     m_modernListbox.multiSelect = IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
' End Sub

' Private Sub OptionSingle_Click()
'     m_modernListbox.multiSelect = IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
' End Sub


' PROPERTIES OF THE USERFORM

' Returns the
Public Function getSelectedItems() As Variant
    getSelectedItems = m_modernListbox.selectedItems
End Function

Public Function getSelectedItem() As Variant
    getSelectedItem = m_modernListbox.SelectedItem
End Function

Public Function getSelectedItemsBrands() As Variant
    getSelectedItemsBrands = m_brandModernListbox.selectedItems
End Function

Public Function getSelectedItemBrand() As Variant
    getSelectedItemBrand = m_brandModernListbox.SelectedItem
End Function

' Prevents the UserForm from being unloaded when X is clicked.
' This is so we can access the selections after
Private Sub UserForm_QueryClose(cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then cancel = True
    ' Hide the Userform and set cancelled to true
    Hide
    
End Sub

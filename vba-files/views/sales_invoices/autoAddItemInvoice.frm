VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} autoAddItemInvoice 
   Caption         =   "Add Items Invoice Automatically"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23130
   OleObjectBlob   =   "autoAddItemInvoice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "autoAddItemInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private WithEvents m_modernListbox As clsModernListbox
Private WithEvents m_brandModernListbox As clsModernListbox

Private m_displayData As Boolean
Private m_brandDisplayData As Boolean

Private Sub brandListboxFrame_Click()

End Sub

Private Sub pgroupListboxFrame_Click()

End Sub

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
    Dim pgroupId As Integer
    pgroupId = CInt(MID(arrayRowToString(arr, 1), 1, 3))
    
    ' Call form.Fill(arr)
    ' form.show
End Sub

Private Sub m_brandModernListbox_ItemSelected(dataRow As Long)
    m_brandDisplayData = True
    If m_brandDisplayData = False Then Exit Sub
    
    Dim form As New formEdit
    Dim arr As Variant
    arr = m_brandModernListbox.GetRow(dataRow)
    Dim brandId As Integer
    brandId = CInt(MID(arrayRowToString(arr, 1), 1, 3))
    ' MsgBox "dataRow=" & dataRow & "=>arrayRowToString= " & arrayRowToString(arr, 1) & ",brandId = " & brandId
    ' Call form.Fill(arr)
    ' form.show
End Sub

Private Sub UserForm_Activate()
    'Generate auto customer code
    ' Dim i As Integer
    ' Dim customerList As Range

    ' Set ws_cus_list = Sheets("CustomerList")
    ' Set customerList = ws_cus_list.Range("customer_list")
    ' LastRow = ws_cus_list.Cells(Rows.Count, "C").End(xlUp).Row + 1

    ' i = 1
    ' Do Until WorksheetFunction.CountA(customerList.Rows(i)) = 0
    '     i = i + 1 'Last row
    ' Loop

    ' With frmCustomer
    '     .customer_code = "CKS" & PadStr(i, 7, "0", xlHAlignRight)
    '     .id = 0
    ' End With
End Sub

Private Sub UserForm_Initialize()
    ' TxtColor forecolor:="#43545F", _
    '     EnterColor:="#005ea2", _
    '     TitleColor:="#ababab", _
    '     AlertColor:="#f72111", _
    '     SuccessColor:="#22ab2b"
    ' tbox.clasBox Me, "Border"
    ' cBtn.classButton Me

    'Define worksheet
    Dim ws_lg, ws_setting, ws_loc, ws_branch, ws_pgroup, ws_brand, ws_add_invoice As Worksheet
    Set ws_loc = ThisWorkbook.Sheets("Location")
    Set ws_lg = ThisWorkbook.Sheets("LocationGroup")
    Set ws_branch = ThisWorkbook.Sheets("Branch")
    Set ws_setting = ThisWorkbook.Sheets("Setting")
    Set ws_pgroup = ThisWorkbook.Sheets("ProductGroup")
    Set ws_brand = ThisWorkbook.Sheets("Brand")
    Set ws_add_invoice = ThisWorkbook.Sheets("AddInvoice")

    'Define range variables
    Dim rngPriceTypeList, rngPriceTypeOption, rngBranch, rngLocationGroup, rngLocation, rngPGroup, rngBrand As Range

    'Define loop variable
    Dim i As Integer
    
    Me.invoice_date = ws_add_invoice.Range("invoice_date").Value
    Me.total_amount.SetFocus

    'Add Dropdown Invoice Price Type
    Me.invoice_price_type.Clear
    Set rngPriceTypeOption = ws_setting.Range("price_type_option") ' Adjust the range as needed
    For i = 1 To rngPriceTypeOption.Rows.Count
        If rngPriceTypeOption.Cells(i) <> "" Then
            Me.invoice_price_type.AddItem rngPriceTypeOption(i)
        End If
    Next i
    Me.invoice_price_type = rngPriceTypeOption.Cells(2) ' Add the first item again at the end

    'Add Dropdown Branch
    Me.branch.Clear
    Set rngBranch = ws_branch.Range("AG7:AG" & ws_branch.Cells(Rows.Count, "C").end(xlUp).row)
    For i = 1 To rngBranch.Rows.Count
        If rngBranch.Cells(i) <> "" Then
            Me.branch.AddItem rngBranch(i)
        End If
    Next i
    Me.branch = rngBranch.Cells(1) ' Add the first item again at the end

    'Add Dropdown Warehouse
    Me.warehouse.Clear
    Set rngLocationGroup = ws_lg.Range("L8:L" & ws_lg.Cells(Rows.Count, "C").end(xlUp).row)
    For i = 1 To rngLocationGroup.Rows.Count
        If rngLocationGroup.Cells(i) <> "" Then
            Me.warehouse.AddItem rngLocationGroup(i)
        End If
    Next i
    Me.warehouse = rngLocationGroup.Cells(1) ' Add the first item again at the end

    'Add Dropdown Location
    Me.location.Clear
    Set rngLocation = ws_loc.Range("O8:O" & ws_loc.Cells(Rows.Count, "C").end(xlUp).row)
    For i = 1 To rngLocation.Rows.Count
        If rngLocation.Cells(i) <> "" Then
            Me.location.AddItem rngLocation(i)
        End If
    Next i
    Me.location = rngLocation.Cells(1) ' Add the first item again at the end

    'product group
    Me.pgroup.Clear
    Set rngPGroup = ws_pgroup.Range("L8:L" & ws_pgroup.Cells(Rows.Count, "C").end(xlUp).row)
    For i = 1 To rngPGroup.Rows.Count
        If rngPGroup.Cells(i) <> "" Then
            Me.pgroup.AddItem rngPGroup(i)
        End If
    Next i

    'product brand
    Me.brand.Clear
    Set rngBrand = ws_brand.Range("J8:J" & ws_brand.Cells(Rows.Count, "C").end(xlUp).row)
    For i = 1 To rngBrand.Rows.Count
        If rngBrand.Cells(i) <> "" Then
            Me.brand.AddItem rngBrand(i)
        End If
    Next i
    Me.brand = rngBrand.Cells(1) ' Add the first item again at the end

      ' OptionMulti.Value = True
    ' CheckBoxHover.Value = True
    ' checkboxScrollbars.Value = False
    ' checkboxSelectOn.Value = True
    pgroupBoxRecords.Value = 15
    brandBoxRecords.Value = 15
    ' checkboxAutoHeight.Value = True
    ' CheckBoxAutoWidth.Value = True

    Set m_modernListbox = Instantiate_clsModernListbox
    Set m_modernListbox.parentFrame = pgroupListboxFrame

    With m_modernListbox
        .HoverOn = True 'CheckBoxHover.Value
        .multiSelect = fmMultiSelectExtended 'IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
        ' .columnWidths = "150;150;90;100;90;90;100"
        .columnWidths = "300"
        .ScrollBars = fmScrollBarsNone 'IIf(checkboxScrollbars.Value = True, fmScrollBarsBoth, fmScrollBarsNone)
        .recordsPerPage = pgroupBoxRecords.Value
        ' .AutomaticHeight = checkboxAutoHeight.Value
        ' .AutomaticWidth = CheckBoxAutoWidth.Value
        ' .HeaderFieldsFromString = Sheets("Movies").Range("A1").Value & ";" & Sheets("Movies").Range("B1").Value & ";" & Sheets("Movies").Range("C1").Value & ";" & Sheets("Movies").Range("D1").Value & ";" & Sheets("Movies").Range("E1").Value & ";" & Sheets("Movies").Range("F1").Value & ";" & Sheets("Movies").Range("G1").Value & ";" & Sheets("Movies").Range("H1").Value
         .HeaderFieldsFromString = ws_pgroup.Range("L7").Value
        Dim rg As Range: Set rg = rngPGroup
        .List = rg.Value
    End With
        
    Set m_brandModernListbox = Instantiate_clsModernListbox
    Set m_brandModernListbox.parentFrame = brandListboxFrame

    With m_brandModernListbox
        .HoverOn = True 'CheckBoxHover.Value
        .multiSelect = fmMultiSelectExtended 'IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
        ' .columnWidths = "150;150;90;100;90;90;100"
        .columnWidths = "300"
        .ScrollBars = fmScrollBarsNone 'IIf(checkboxScrollbars.Value = True, fmScrollBarsBoth, fmScrollBarsNone)
        .recordsPerPage = brandBoxRecords.Value
        ' .AutomaticHeight = checkboxAutoHeight.Value
        ' .AutomaticWidth = CheckBoxAutoWidth.Value
        ' .HeaderFieldsFromString = Sheets("Movies").Range("A1").Value & ";" & Sheets("Movies").Range("B1").Value & ";" & Sheets("Movies").Range("C1").Value & ";" & Sheets("Movies").Range("D1").Value & ";" & Sheets("Movies").Range("E1").Value & ";" & Sheets("Movies").Range("F1").Value & ";" & Sheets("Movies").Range("G1").Value & ";" & Sheets("Movies").Range("H1").Value
         .HeaderFieldsFromString = ws_brand.Range("J7").Value
        Dim rgBrand As Range: Set rgBrand = rngBrand
        .List = rgBrand.Value
    End With

    ' Show controls
    Call ShowControlSection
    
    ' Set Userform to screen size


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
Private Sub pgroupBoxRecords_Change()
    ' If Len(Trim(pgroupBoxRecords.Value)) = 0 Then Exit Sub
    ' m_modernListbox.recordsPerPage = Trim(pgroupBoxRecords.Value)
End Sub

Private Sub brandBoxRecords_Change()
    ' If Len(Trim(brandBoxRecords.Value)) = 0 Then Exit Sub
    ' m_brandModernListbox.recordsPerPage = Trim(brandBoxRecords.Value)
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


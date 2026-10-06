Attribute VB_Name = "Main"
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM

' Run this to see the Listbox demo
Sub ListBoxDemo()
    
    ' Display the UserForm
    Dim f As New UserFormClass
    f.Show
    
    ' Retrieve the array of selected items
    Dim selectedData As Variant
    selectedData = f.getSelectedItems()
    If IsEmpty(selectedData) = True Then Exit Sub
    
    ' Write selected items to the worksheet
    Sheets("Output").Range("A1").CurrentRegion.Offset(1).ClearContents
    Call ArrayToRange(selectedData, Sheets("Output").Range("A2"))
    
    Dim slectedDataBrands As Variant
    slectedDataBrands = f.getSelectedItemsBrands()
    If IsEmpty(slectedDataBrands) = True Then Exit Sub

    Sheets("Output").Range("B1").CurrentRegion.Offset(1).ClearContents
    Call ArrayToRange(slectedDataBrands, Sheets("Output").Range("B2"))
End Sub

Public  Sub listBoxAutoAddItemInvoice()
    Dim frm As New autoAddItemInvoice
    frm.Show
    ' Retrieve the array of selected items
    Dim selectedData As Variant
    selectedData = frm.getSelectedItems()
    If IsEmpty(selectedData) = True Then Exit Sub
    
    ' Write selected items to the worksheet
    Sheets("Output").Range("A1").CurrentRegion.Offset(1).ClearContents
    Call ArrayToRange(selectedData, Sheets("Output").Range("A2"))
    
    Dim slectedDataBrands As Variant
    slectedDataBrands = frm.getSelectedItemsBrands()
    If IsEmpty(slectedDataBrands) = True Then Exit Sub

    Sheets("Output").Range("B1").CurrentRegion.Offset(1).ClearContents
    Call ArrayToRange(slectedDataBrands, Sheets("Output").Range("B2"))

End Sub
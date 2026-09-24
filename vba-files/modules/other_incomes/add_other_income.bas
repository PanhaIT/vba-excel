Attribute VB_Name = "add_other_income"
Option Explicit

' oic_total_amount_paid
' oic_receipt_code
' oic_amount_due
' oic_total_balance
' oic_option_income_type
' oic_customer_name
' oic_recipient_option
' oic_option_income_from
' oic_status_note

' oic_description
' oic_exchange_rate
' oic_current_date
' oic_status
' oic_amount_kh_paid
' oic_amount_en_paid
' oic_invoice_no
' oic_customer_code
' oic_customer_id
' oic_receipt_no

Sub SaveDataOtherIncome()
    Dim paidDate As Date
    Dim sheetNo, statusNote, receiptNo, customerName, incomeType,description, branchName As String
    Dim invoiceNo, incomeId, customerId, incomeTypeId, branchId, saleId,invoiceWeek,invoiceMonth,invoiceYear As Integer
    Dim totalAmountDue, totalAmountPaidKh, totalAmountPaidUsd, totalAmountPaid, totalBalance As Double
    Dim answer As VbMsgBoxResult
    Dim wb As Workbook
    Dim ws_oic, ws As Worksheet

    Set wb = ActiveWorkbook
    Set ws_oic = wb.Sheets("OtherIncome")
    answer = MsgBox("Are you want to save other income ?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Receive Payment")

    receiptNo = ws_oic.Range("oic_receipt_code").Value
    sheetNo = ws_oic.Range("oic_invoice_sheet_no").Value
    invoiceNo = ws_oic.Range("oic_invoice_no").Value
    paidDate = ws_oic.Range("oic_current_date").Value
    customerName = ws_oic.Range("oic_customer_name").Value

    incomeType = ws_oic.Range("oic_option_income_from").Value
    branchName = ws_oic.Range("oic_option_income_type").Value

    customerId = ws_oic.Range("oic_customer_id").Value
    incomeTypeId = ws_oic.Range("oic_income_id").Value
    branchId = MID(ws_oic.Range("oic_option_income_type").Value, 1, 2)
    saleId = ws_oic.Range("oic_recipient_option").Value

    description = ws_oic.Range("oic_description").Value
    invoiceWeek = ws_oic.Range("oic_week").Value
    invoiceMonth = month(ws_oic.Range("oic_current_date").Value)
    invoiceYear = Year(ws_oic.Range("oic_current_date").Value)

    totalAmountDue = ws_oic.Range("oic_total_amount_paid").Value       'ws_oic.Range("oic_amount_due").Value
    totalAmountPaidKh = Sheets("OtherIncome").Range("oic_amount_kh_paid").Value
    totalAmountPaidUsd = Sheets("OtherIncome").Range("oic_amount_en_paid").Value
    totalAmountPaid      = ws_oic.Range("oic_total_amount_paid").Value
    totalBalance         = ws_oic.Range("oic_total_balance").Value

    totalBalance = 0
    ' MsgBox "totalAmountPaidKh=" & totalAmountPaidKh & ", totalAmountPaidUsd=" & totalAmountPaidUsd & ", totalAmountPaid=" & totalAmountPaid & ", totalBalance=" & totalBalance
    ' Exit Sub

    ' Run resetWorkSheetsFilter()
    ' If totalAmountPaid > totalAmountDue Then
    '     MsgBox "Total amount paid is bigger than total amount due, please try again."
    '     Exit Sub
    ' ElseIf (totalAmountPaid + totalDiscountKh) > totalAmountDue Then
    '     MsgBox "Total amount paid and discount are bigger than total amount due, please try again."
    '     Exit Sub
    ' End If

    ' Sales Invoice Status=> -1 = Edit, 1  = issue, 2  = fullfilied ,3  = partial
    If ActiveSheet.Range("oic_status").Value = 2 Then
        MsgBox "OtherIncome already saved."
        Exit Sub
    ElseIf saleId = "" Or ActiveSheet.Range("oic_module_types_option").Value = "" Or ActiveSheet.Range("oic_current_date").Value = "" Or ActiveSheet.Range("oic_option_income_from").Value = "" Or ActiveSheet.Range("oic_option_income_type").Value = "" And (ActiveSheet.Range("oic_amount_en_paid") = 0 Or ActiveSheet.Range("oic_amount_kh_paid").Value = 0) Then
        MsgBox "Please select or input require field, try again."
        Exit Sub
    Else
        If answer = vbYes Then
            '*******Start save amount paid*******
            If totalAmountPaid > 0 Then
                Dim i As Long
                Dim ic_data As Range 'ic_data : income date

                i = 1
                Set ic_data = Sheets("IncomeData").Range("B5:W5")
                Do Until WorksheetFunction.CountA(ic_data.Rows(i)) = 0
                    i = i + 1
                Loop

                ic_data.Cells(i,1)  = ""
                ic_data.Cells(i,2)  = paidDate 'invoice date
                ic_data.Cells(i,3)  = "019" 'module type id
                ic_data.Cells(i,4)  = branchId 'branch id
                ic_data.Cells(i,5)  = incomeTypeId 'income type id
                ic_data.Cells(i,6)  = saleId 'sale id
                ic_data.Cells(i,7)  = customerId 'customer id
                ic_data.Cells(i,8)  = paidDate 'paid date
                ic_data.Cells(i,9)  = receiptNo 'receipt no.
                ic_data.Cells(i,10) = branchName 'branch name
                ic_data.Cells(i,11) = incomeType 'income type
                ic_data.Cells(i,12) = description
                ic_data.Cells(i,13) = customerName 'customer name

                ic_data.Cells(i,14) = totalAmountDue 'amount due
                ic_data.Cells(i,15) = totalAmountPaidKh 'amount paid riel
                ic_data.Cells(i,16) = totalAmountPaidUsd 'amount paid dollar
                ic_data.Cells(i,17) = totalAmountPaid 'total amount
                ic_data.Cells(i,18) = 0 'discount
                ic_data.Cells(i,19) = totalBalance 'balance

                ic_data.Cells(i,20) = invoiceWeek 'week
                ic_data.Cells(i,21) = invoiceMonth 'month
                ic_data.Cells(i,22) = invoiceYear 'year
            End If
            '*******End save amount paid*******

            Call clearOtherIncome
            Sheets("IncomeData").select
            incomeId = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
            ActiveSheet.Range("B" & incomeId & ":" & "W" & incomeId).Select
            MsgBox "Other Income successful saved."
        Else
            exit sub
        End If
    End If
End Sub

Function clearOtherIncome()
    ActiveSheet.Range("oic_current_date").Value = ""
    ActiveSheet.Range("oic_invoice_no").Value = ""
    ActiveSheet.Range("oic_amount_kh_paid").Value = ""
    ActiveSheet.Range("oic_amount_en_paid").Value = ""
    ' ActiveSheet.Range("oic_discount").Value = ""
    'ActiveSheet.Range("oic_recipient_option").Value = ""
    ActiveSheet.Range("oic_receipt_no").Value = ActiveSheet.Range("oic_receipt_no").Value + 1
End Function

Sub copyOtherIncome()
    Sheets("OtherIncome").Range("Print_Area").Copy
End Sub

Sub currentDateOic()
    ActiveSheet.Range("oic_current_date").Value = Format(Now(), "mm/dd/yyyy")
End Sub

Sub ExchangeRateOic()
    ActiveSheet.Range("oic_exchange_rate").Value = Sheets("NBC_Exchange_Rate").Range("ExchangeRateToday").Value
End Sub
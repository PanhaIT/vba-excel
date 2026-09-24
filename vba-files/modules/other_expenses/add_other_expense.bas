Attribute VB_Name = "add_other_expense"

Sub SaveOtherExpense()
    Dim expenseCode,expenseNote,referenceInvoiceNo,moduleTypeId,branchId,vendorId,expenseTypeId,branchName,expenseType,expenseWeek,expenseMonth,expenseYear,lastExpId As String
    Dim indexColor As Integer
    Dim totalAmountPaidKh,totalAmountPaidUsd AS Double
    Dim paidDate AS date
    Dim wb As Workbook
    Dim exp As Worksheet
    Dim answer AS VbMsgBoxResult

    set exp = Sheets("OtherExpense")
    answer = MsgBox("Are you want to save expense " & invoiceSheet & " ?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Other Expense")

    expenseCode          = exp.Range("exp_code").Value
    referenceInvoiceNo   = exp.Range("c_exp_ref_invoice_no").Value
    paidDate             = exp.Range("c_exp_current_date").Value
    expenseType          = exp.Range("other_expense_type_option").Value
    branchName           = exp.Range("other_exp_branch_option").Value
    expenseNote          = exp.Range("c_exp_description").Value
    moduleTypeId         = MID(exp.Range("c_exp_module_types_option").Value,1,2)
    expenseTypeId        = exp.Range("c_expense_type_id").Value
    branchId             = exp.Range("c_exp_branch_id").Value
    vendorId             = exp.Range("c_exp_vendor_id").Value
    memberId             = exp.Range("exp_member_id").Value   
    expenseNote          = exp.Range("c_rp_description").Value
    expenseWeek          = exp.Range("exp_week").Value
    expenseMonth         = MONTH(exp.Range("c_exp_current_date").Value)
    expenseYear          = YEAR(exp.Range("c_exp_current_date").Value)

    totalAmountPaidKh    = exp.Range("c_exp_amount_kh_paid").Value
    totalAmountPaidUsd   = exp.Range("c_exp_amount_en_paid").Value

    Run resetWorksheetsFilter()

    If ActiveSheet.Range("c_exp_branch_id").Value = "" Or ActiveSheet.Range("c_expense_type_id").Value = "" Or ActiveSheet.Range("c_exp_vendor_id").Value = "" Or ActiveSheet.Range("c_exp_module_types_option").Value = "" Or ActiveSheet.Range("c_exp_current_date").Value = "" And (ActiveSheet.Range("c_exp_amount_kh_paid").Value = 0 Or ActiveSheet.Range("c_exp_amount_en_paid").Value = 0) Then
        MsgBox "Please select or input require field, try again."
        Exit Sub
    Else
        If totalAmountPaidKh > 0 Or totalAmountPaidUsd >0 Then
            If answer = vbYes Then
                Dim k As Long
                Dim oex_data As Range

                k = 1
                Set oex_data = Sheets("ExpenseData").Range("B5:S5")
                Do Until WorksheetFunction.CountA(oex_data.Rows(k)) = 0
                    k = k + 1
                Loop
                oex_data.Cells(k,1)  = expenseCode 'invoice no.
                oex_data.Cells(k,2)  = paidDate 'Paid date
                oex_data.Cells(k,3)  = moduleTypeId 'module type id
                oex_data.Cells(k,4)  = branchId 'branch id
                oex_data.Cells(k,5)  = expenseTypeId 'expense type id
                oex_data.Cells(k,6)  = "" 'customer id
                oex_data.Cells(k,7)  = vendorId 'vendor id
                oex_data.Cells(k,8)  = " " 'member id
                oex_data.Cells(k,9)  = branchName 'branch name
                oex_data.Cells(k,10) = expenseType 'expense type name
                oex_data.Cells(k,11) = expenseNote 'description
                oex_data.Cells(k,12) = referenceInvoiceNo 'reference invoice code
                oex_data.Cells(k,13) = totalAmountPaidKh 'amount paid kh
                oex_data.Cells(k,14) = totalAmountPaidUsd 'amount paid usd
                oex_data.Cells(k,15) = expenseNote 'remark

                oex_data.Cells(k,16) = expenseWeek 'week
                oex_data.Cells(k,17) = expenseMonth 'month
                oex_data.Cells(k,18) = expenseYear 'year

                clearOtherExpense (obj)
                Sheets("ExpenseData").select
                lastExpId = Sheets("ExpenseData").Cells(Rows.Count, "B").End(xlUp).Row
                ActiveSheet.Range("B" & lastExpId & ":" & "S" & lastExpId).Select
                MsgBox "Expense successful saved."
                exit sub
            Else
                exit sub
            End If
        Else
            MsgBox "Amount is zero, please input amount before save."
            exit sub
        End If
    End If
End Sub

Sub checkExistExpenseAmount()
    Dim totalAmountPaidKh,totalAmountPaidUSD,totalAmount,amountKhCheck,amountEnCheck,totalAmountCheck AS Double
    Dim exchangeRate, status, checkYear, checkMonth, expenseYear AS Integer
    Dim LastRowAmountKh,i AS Long
    Dim payDate,payMonth AS string
    
    totalAmountPaidKh  = ActiveSheet.Range("c_exp_amount_kh_paid").value
    totalAmountPaidUSD = ActiveSheet.Range("c_exp_amount_en_paid").value
    exchangeRate       = 4000
    amountKhCheck      = 0
    amountEnCheck      = 0
    totalAmount        = totalAmountPaidKh + totalAmountPaidUSD * exchangeRate
    LastRowAmountKh    = Sheets("ExpenseData").Cells(Rows.Count, "M").End(xlUp).Row
    expenseYear        = ActiveSheet.Range("exp_year").value
    payDate            = ActiveSheet.Range("c_exp_current_date").Value
    payMonth           = MONTH(payDate) 

    If totalAmount>0  Then
        ActiveSheet.Range("c_other_expense_status").value = 0
        For i = 5 To LastRowAmountKh
            amountKhCheck    = Sheets("ExpenseData").Cells(i, 14)
            amountEnCheck    = Sheets("ExpenseData").Cells(i, 15)
            checkMonth       = Sheets("ExpenseData").Cells(i, 18)
            checkYear        = Sheets("ExpenseData").Cells(i, 19)

            totalAmountCheck = amountKhCheck + amountEnCheck * exchangeRate

            If Sheets("ExpenseData").Cells(i, 14).Value = totalAmountPaidKh and totalAmountPaidKh > 0  and expenseYear = checkYear AND payMonth = checkMonth Then
                status = 1
                ActiveSheet.Range("c_other_expense_status").value = 1
                ActiveSheet.Range("exp_alert_to_check_amount").value = "Amount riel is exist, please check transaction before save."
                exit sub
            ElseIf Sheets("ExpenseData").Cells(i, 15).Value = totalAmountPaidUSD and totalAmountPaidUSD > 0 and expenseYear = checkYear AND payMonth = checkMonth Then
                status = 1
                ActiveSheet.Range("c_other_expense_status").value = 1
                ActiveSheet.Range("exp_alert_to_check_amount").value = "Amount dollar is exist, please check transaction before save."
                exit sub
            ElseIf totalAmount = totalAmountCheck and totalAmountCheck >0 and expenseYear = checkYear AND payMonth = checkMonth Then
                status = 1
                ActiveSheet.Range("c_other_expense_status").value = 1
                ActiveSheet.Range("exp_alert_to_check_amount").value = "Amount riel/dollar is exist, please check transaction before save."
                exit sub
            End If
        Next i
        If ActiveSheet.Range("c_other_expense_status").value = 0 Then
            ActiveSheet.Range("exp_alert_to_check_amount").value = ""
        End If
    Else
        ActiveSheet.Range("c_other_expense_status").value = 0
        MsgBox "Please check require field, then try again."
        exit sub
    End If
End Sub

Function clearOtherExpense(obj)
    ActiveSheet.Range("c_exp_ref_invoice_no").Value = "" 
    ActiveSheet.Range("c_other_expense_status").Value = 0
    ActiveSheet.Range("exp_member_option").Value = ""
    ActiveSheet.Range("c_exp_current_date").Value = ""
    ActiveSheet.Range("c_exp_amount_kh_paid").Value = ""
    ActiveSheet.Range("c_exp_amount_en_paid").Value = ""
    ActiveSheet.Range("c_exp_qrcode").Value = ""
    ActiveSheet.Range("c_other_expense_no").Value = ActiveSheet.Range("c_rp_receipt_no").Value + 1
End Function

Sub CurrentDateOtherExpense()
    ActiveSheet.Range("c_exp_current_date").Value = Format(Now(), "mm/dd/yyyy")
End Sub

Sub ExchangeRateOtherExpense()
    ActiveSheet.Range("c_exp_exchange_rate").Value = Sheets("NBC_Exchange_Rate").Range("ExchangeRateToday").Value
End Sub

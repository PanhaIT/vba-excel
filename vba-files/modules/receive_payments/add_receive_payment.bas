Attribute VB_Name = "add_receive_payment"

'Receive Payment
' 1.id
' 2.date
' 3.company_id
' 4.branch_id
' 5.cgroup_id
' 6.customer_id
' 7.deposit_to
' 8.reference
' 9.remark
' 10.created
' 11.created_by
' 12.is_active

' Receive Payment Detail
' 1.id
' 2.receive_payment_id
' 3.invoice_code
' 4.sales_invoice_id
' 5.sales_invoice_receipt_id
' 6.general_ledger_detail_id
' 7.Exchange Rate ID
' 8.Exchange Rate
' 9.amount_due
' 10.paid
' 11.paid_other
' 12.discount
' 13.discount_percent
' 14.discount_other
' 15.balance
' 16.due_date
' 17.is_void

' Sales Invoice Receipt
' 1.id
' 2.sales_invoice_id
' 3.invoice_code
' 4.branch_id
' 5.exchange_rate_id
' 6.currency_center_id
' 7.chart_account_id
' 8.receipt_code
' 9.Exchange Rate
' 10.amount_due
' 11.amount_due_other
' 12.amount_paid_us
' 13.amount_paid_other
' 14.discount_percent
' 15.discount_us
' 16.discount_other
' 17.balance
' 18.balance_other
' 19.change
' 20.change_other
' 21.pay_date
' 22.due_date
' 23.created
' 24.created_by
' 25.is_void

' Sales Invoice Data
' 1.Id
' 2.Tax Inv Code
' 3.Invoice No
' 4.Invoice Date
' 5.Created Date
' 6.Branch ID
' 7.Location Group ID
' 8.Location ID
' 9.Price Type ID
' 10.Discount Type ID
' 11.Sale ID
' 12.Customer ID
' 13.Payment Term Id
' 14.Currency Id
' 15.Chart Account AR
' 16.Sale Name
' 17.Customer Name
' 18.Exchange Rate
' 19.Vat Chart Account Id
' 20.Vat Percent
' 21.Vat Setting Id
' 22.Vat Cal
' 23.Total Amount
' 24.Total Vat
' 25.Discount
' 26.Total Deposit
' 27.Balance
' 28.Status
' 29.Note
' 30.Invoice Type
' 31.Week
' 32.Month
' 33.Year
' 34.Invoice ID

Sub SaveReceivePayment()
    OnStart
    'Loop variable
    Dim LastRowSaleInvoiceDate As Long

    'Workbook, worksheet Variable
    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_gl,ws_gld,ws_cus,ws_inv_receipt,ws_rp_data,ws_rp_detail,ws_file_setting,ws_rp,ws_inv As Worksheet

    'Receive Payment Variable
    Dim sheetNo,statusNote,receiptNo,invoiceCode,branchName As String
    Dim gl_id,gld_id,rpCurrencyId,invoiceId, sirID, rpdID, cgroupId,chartAccountId, rpID, indexColor, statusInvoice As Integer
    Dim coaDefaultRange,ws_gl_data,ws_gld_data,rng_customer,rng_rp,rng_rpd,rng_sir As Range
    Dim createdDate,invoiceDueDate As Date
    Dim gld_rp_dis_id,gld_rp_id,locationId,locationGroupId As Integer
    DIm coaPaymentDiscount , coaReceivable, coaCashAndBank As Integer
    
    'Amount Variable
    Dim grandTotalPaid,totalDiscount,totalAmountPaidEn,totalAmountPaid,totalAmountPaidBefore,totalBalance AS Double
    Dim exchangeRate,totalDiscounEn,totalAmountDue,totalDiscounKh,totalLastDiscountEn,totalLastDiscountKh,totalAmountPaidKh,totalChangeKh,totalChangeEn As Double

    'Setting Variable
    Dim databasePath, databaseFile, dataInputPath As String
    Dim netDays As Integer
    
    Dim confirmBoxAlert AS VbMsgBoxResult

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
    Set ws_rp           = activeWorkbook.Sheets("AddReceivePayment")
    Set ws_cus          = activeWorkbook.Sheets("CustomerList")
    Set coaDefaultRange = activeWorkbook.Sheets("ChartAccountDefault").Range("chart_account_default_list")

    coaPaymentDiscount  = Application.VLookup(20, coaDefaultRange, 5, False) 'Payment Discount
    coaReceivable       = Application.VLookup(14, coaDefaultRange, 5, False) 'Accounts Receivable
    coaCashAndBank      = Application.VLookup(19, coaDefaultRange, 5, False) 'Cash And Bank
    'debug.Print " , coaPaymentDiscount=" & coaPaymentDiscount & " , coaReceivable=" & coaReceivable & " , coaCashAndBank=" & coaCashAndBank

    confirmBoxAlert       = MsgBox("Are you want to save this payment?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "Receive Payment")
    receiptNo             = ws_rp.Range("rp_receipt_code").Value
    invoiceId             = ws_rp.Range("rp_invoice_id").Value
    invoiceCode           = ws_rp.Range("rp_invoice_code").Value
    invoiceDate           = ws_rp.Range("rp_invoice_date").Value
    paidDate              = ws_rp.Range("rp_current_date").Value
    customerName          = ws_rp.Range("rp_customer_name").Value
    branchName            = MID(ws_rp.Range("rp_branch").Value,4,200)
    chartAccountId        = ws_rp.Range("rp_chart_account_id").Value
    customerId            = ws_rp.Range("rp_customer_id").Value
    rpCurrencyId          = ws_rp.Range("rp_currency_id").Value
    branchId              = CInt(MID(ws_rp.Range("rp_branch").Value,1,2))
    saleId                = ws_rp.Range("rp_recipient_option").Value
    netDays               = ws_rp.Range("rp_payment_net_days").Value
    locationGroupId       = ws_rp.Range("rp_location_group_id").Value
    locationId            = ws_rp.Range("rp_location_id").Value

    invoiceDueDate        = 0
    If invoiceDate > 0 Then
        invoiceDueDate    = DateAdd("d", netDays, invoiceDate)
    End If

    rpDescription         = ws_rp.Range("rp_description").Value
    invoiceWeek           = ws_rp.Range("rp_week").Value
    invoiceMonth          = MONTH(ws_rp.Range("rp_current_date").Value)
    invoiceYear           = YEAR(ws_rp.Range("rp_current_date").Value)

    If ( ws_rp.Range("rp_exchange_rate").Value > 0) Then 
        exchangeRate          = ws_rp.Range("rp_exchange_rate").Value
    Else
        exchangeRate          = 4000
    End If

    totalAmountDue        = ws_rp.Range("rp_amount_due").Value
    totalAmountPaid       = ws_rp.Range("rp_total_paid").Value
    totalAmountPaidBefore = ws_rp.Range("rp_amount_paid_before").Value
    If (totalAmountPaid > totalAmountDue) Then 
        If (rpCurrencyId = 1) Then 'Dollar
            totalAmountPaidKh     = 0
            totalAmountPaidEn     = totalAmountDue
        ElseIf (rpCurrencyId = 2) Then 'Riel
            totalAmountPaidKh     = totalAmountDue
            totalAmountPaidEn     = 0
        End If
        totalAmountPaid       = totalAmountDue
    Else
        totalAmountPaidKh     = ws_rp.Range("rp_amount_paid_kh").Value
        totalAmountPaidEn     = ws_rp.Range("rp_amount_paid_en").Value
    End If

    totalDiscounKh        = ws_rp.Range("rp_discount_kh").Value
    totalDiscounEn        = ws_rp.Range("rp_discount_en").Value

    If (rpCurrencyId = 2) Then 'Dollar
        totalDiscount     = totalDiscounEn * exchangeRate + totalDiscounKh
    Else
        totalDiscount     = totalDiscounEn + totalDiscounKh / exchangeRate
    End If

    grandTotalPaid        = totalAmountPaid + totalDiscount
    totalChangeKh         = 0 'ws_rp.Range("rp_amount_change_kh").Value
    totalChangeEn         = 0 'ws_rp.Range("rp_amount_change_en").Value
    totalBalance          = ws_rp.Range("rp_total_balance").Value 'rp_total_balance
    createdDate           = Format(Now(), "yyyy/mm/dd hh:mm:ss")
    Run resetAllFilter()

    If totalDiscount > totalAmountDue Then
        MsgBox "Total discount is bigger than total amount due, please try again."
        Exit Sub
    End If

    'Sales Invoice Status=> -1 = Edit, 1  = issue, 2  = fullfilied ,3  = partial
    If totalBalance = 0 Then
        statusInvoice = 2
        indexColor    = 43
    ElseIf totalAmountDue > grandTotalPaid and grandTotalPaid > 0 Then
        statusInvoice = 3
        indexColor    = 44
    ElseIf totalBalance = totalAmountDue Then
        statusInvoice = 1
        indexColor    = 3
    Else
        ' statusInvoice = -1
        ' indexColor    = 1
        Exit Sub
    End If

    If ActiveSheet.Range("rp_status").Value = 2 Then
        MsgBox "Receive payment already saved."
        Exit Sub
    ElseIf saleId = "" Or ActiveSheet.Range("rp_module_types_option").Value = "" Or ActiveSheet.Range("rp_current_date").Value = "" Or ActiveSheet.Range("rp_invoice_no").Value = "" And (ActiveSheet.Range("rp_amount_paid_en") = 0 Or ActiveSheet.Range("rp_amount_paid_kh").Value = 0) Then
        MsgBox "Please select or input require field, try again."
        Exit Sub
    Else
        If confirmBoxAlert = vbYes Then
            ' ***************Start save amount paid***************
            If totalAmountPaid > 0 Then
                databasePath        = ws_file_setting.Range("database_path").Value
                databaseFile        = databasePath & "database_kscm.xlsm"
                dataInputPath       = ws_file_setting.Range("data_input_path").value

                If (IsWorkbookOpen(CStr(databaseFile)) = FALSE) Then
                    Set wbDatabase  = Workbooks.Open(databaseFile)  'add here the path of your text file
                End If

                Set ws_inv          = wbDatabase.Sheets("SaleInvoiceData")
                Set ws_inv_receipt  = wbDatabase.Sheets("SalesInvoiceReceipt")
                Set ws_rp_data      = wbDatabase.Sheets("ReceivePayment")
                Set ws_rp_detail    = wbDatabase.Sheets("ReceivePaymentDetail")
                Set rng_customer    = ws_cus.Range("C5:P" & ws_cus.Cells(Rows.Count,"C").End(xlUp).Row)
                Set rng_rp          = ws_rp_data.Range("C8:N8")
                Set rng_rpd         = ws_rp_detail.Range("C8:S8")
                Set rng_sir         = ws_inv_receipt.Range("C8:AA8")
                
                Set ws_gl           = wbDatabase.sheets("GeneralLedger")
                Set ws_gld          = wbDatabase.Sheets("GeneralLedgerDetail")
                Set ws_gl_data      = ws_gl.Range("C8:Al8")
                Set ws_gld_data     = ws_gld.Range("C8:AC8")

                gl_id               = ws_gl.Cells(Rows.Count, "C").End(xlUp).Row - 6 'gl = general ledger
                gld_id              = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6 'gld = general ledger detail

                cgroupId            = CInt(MID(Application.VLOOKUP(customerId,rng_customer,4,False),1,3)) 
                rpID                = ws_rp_data.Cells(Rows.Count, "C").End(xlUp).Row - 6
                rpdID               = ws_rp_detail.Cells(Rows.Count, "C").End(xlUp).Row - 6
                sirID               = ws_inv_receipt.Cells(Rows.Count,"C").End(xlUp).Row - 6

                ''***************Insert Receive Payment***************
                rng_rp.Cells(rpID,1)  = rpID 'id
                rng_rp.Cells(rpID,2)  = paidDate 'receive payment date
                rng_rp.Cells(rpID,3)  = 1 'company id
                rng_rp.Cells(rpID,4)  = branchId 'branch id
                rng_rp.Cells(rpID,5)  = cgroupId 'cgroup id
                rng_rp.Cells(rpID,6)  = customerId 'customer id
                rng_rp.Cells(rpID,7)  = coaCashAndBank 'chart account cash and bank
                rng_rp.Cells(rpID,8)  = receiptNo 'reference
                rng_rp.Cells(rpID,9)  = rpDescription 'remark
                rng_rp.Cells(rpID,10) = createdDate 'created
                rng_rp.Cells(rpID,11) = saleId 'created by 
                rng_rp.Cells(rpID,12) = 1 'is active
                rng_rp.Cells(rpID,13) = invoiceWeek 'week
                rng_rp.Cells(rpID,14) = invoiceMonth 'month
                rng_rp.Cells(rpID,15) = invoiceYear 'year 

                If (rng_rp.Cells(rpID,1) = rpID) Then
                    ''***************Insert Sales Invoice Receipt***************
                    rng_sir.Cells(sirID, 1)  = sirID 'id
                    rng_sir.Cells(sirID, 2)  = invoiceId 'sales_invoice_id
                    rng_sir.Cells(sirID, 3)  = invoiceCode 'invoice_code
                    rng_sir.Cells(sirID, 4)  = branchId 'branch_id
                    rng_sir.Cells(sirID, 5)  = "" 'exchange_rate_id
                    rng_sir.Cells(sirID, 6)  = rpCurrencyId 'currency_center_id
                    rng_sir.Cells(sirID, 7)  = coaCashAndBank 'chart_account_id
                    rng_sir.Cells(sirID, 8)  = receiptNo 'receipt_code
                    rng_sir.Cells(sirID, 9)  = exchangeRate 'Exchange Rate
                    rng_sir.Cells(sirID, 10) = totalAmountDue 'amount_due
                    rng_sir.Cells(sirID, 11) = 0 'amount_due_other
                    rng_sir.Cells(sirID, 12) = totalAmountPaidEn 'amount_paid_us
                    rng_sir.Cells(sirID, 13) = totalAmountPaidKh 'amount_paid_other
                    rng_sir.Cells(sirID, 14) = 0 'discount_percent
                    rng_sir.Cells(sirID, 15) = totalDiscounEn 'discount_us
                    rng_sir.Cells(sirID, 16) = totalDiscounKh 'discount_other
                    rng_sir.Cells(sirID, 17) = totalBalance 'balance
                    rng_sir.Cells(sirID, 18) = 0 'balance_other
                    rng_sir.Cells(sirID, 19) = totalChangeEn 'change
                    rng_sir.Cells(sirID, 20) = totalChangeKh 'change_other
                    rng_sir.Cells(sirID, 21) = paidDate 'pay_date
                    If (invoiceDueDate > 0) Then 
                    rng_sir.Cells(sirID, 22) = invoiceDueDate 'due_date
                    Else
                    rng_sir.Cells(sirID, 22) = "" 'due_date
                    End If
                    rng_sir.Cells(sirID, 23) = createdDate 'created
                    rng_sir.Cells(sirID, 24) = saleId 'created_by
                    rng_sir.Cells(sirID, 25) = 0 'is_void
                    rng_sir.Cells(sirID, 26) = invoiceWeek 'week
                    rng_sir.Cells(sirID, 27) = invoiceMonth 'month
                    rng_sir.Cells(sirID, 28) = invoiceYear 'year 

                    '***************Insert Receive Payment Detail***************
                    If (rng_sir.Cells(sirID, 1) = sirID) Then
                        rng_rpd.Cells(rpdID, 1)  = rpdID 'id
                        rng_rpd.Cells(rpdID, 2)  = rpID 'receive_payment_id
                        rng_rpd.Cells(rpdID, 3)  = invoiceCode 'invoice_code
                        rng_rpd.Cells(rpdID, 4)  = invoiceId 'sales_invoice_id
                        rng_rpd.Cells(rpdID, 5)  = sirID 'sales_invoice_receipt_id
                        rng_rpd.Cells(rpdID, 6)  = "" 'general_ledger_detail_id
                        rng_rpd.Cells(rpdID, 7)  = "" 'Exchange Rate ID
                        rng_rpd.Cells(rpdID, 8)  = exchangeRate 'Exchange Rate
                        rng_rpd.Cells(rpdID, 9)  = totalAmountDue 'amount_due
                        rng_rpd.Cells(rpdID, 10) = totalAmountPaidEn 'paid
                        rng_rpd.Cells(rpdID, 11) = totalAmountPaidKh 'paid_other
                        rng_rpd.Cells(rpdID, 12) = 0 'discount_percent
                        rng_rpd.Cells(rpdID, 13) = totalDiscounEn 'discount
                        rng_rpd.Cells(rpdID, 14) = totalDiscounKh 'discount_other
                        rng_rpd.Cells(rpdID, 15) = totalBalance 'balance
                        If (invoiceDueDate > 0) Then 
                        rng_rpd.Cells(rpdID, 16) = invoiceDueDate 'due_date
                        Else
                        rng_rpd.Cells(rpdID, 16) = "" 'due_date
                        End If
                        rng_rpd.Cells(rpdID, 17) = 0 'is_void
                        rng_rpd.Cells(rpdID, 18) = invoiceWeek 'week
                        rng_rpd.Cells(rpdID, 19) = invoiceMonth 'month
                        rng_rpd.Cells(rpdID, 20) = invoiceYear 'year 
                    End If

                    ''***************insert general ledger***************
                    ws_gl_data.Cells(gl_id,1)  = gl_id 'id
                    ws_gl_data.Cells(gl_id,3)  = invoiceId 'sales_invoice_id
                    ws_gl_data.Cells(gl_id,4)  = sirID 'sales_invoice_receipt_id
                    ws_gl_data.Cells(gl_id,5)  = rpID 'receive_payment_id
                    ws_gl_data.Cells(gl_id,21) = customerId 'receive_from_id
                    ws_gl_data.Cells(gl_id,22) = customerName 'receive_from_name
                    ws_gl_data.Cells(gl_id,23) = paidDate 'date
                    ws_gl_data.Cells(gl_id,24) = receiptNo 'reference
                    ws_gl_data.Cells(gl_id,25) = 0 'total_deposit
                    ws_gl_data.Cells(gl_id,27) = createdDate 'created
                    ws_gl_data.Cells(gl_id,28) = saleId 'created_by
                    ws_gl_data.Cells(gl_id,29) = 1 'is_sys
                    ws_gl_data.Cells(gl_id,30) = 0 'is_adj
                    ws_gl_data.Cells(gl_id,31) = 1 'is_approve
                    ws_gl_data.Cells(gl_id,32) = 0 'is_depreciated
                    ws_gl_data.Cells(gl_id,33) = 0 'is_retained_earnings
                    ws_gl_data.Cells(gl_id,34) = 0 'deposit_type
                    ws_gl_data.Cells(gl_id,35) = 1 'is_active
                    ws_gl_data.Cells(gl_id,36) = invoiceWeek 'week
                    ws_gl_data.Cells(gl_id,37) = invoiceMonth 'month
                    ws_gl_data.Cells(gl_id,38) = invoiceYear 'year

                    ''***************insert general ledger detail***************
                    If (totalAmountPaid > 0) Then
                        gld_rp_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                        ' debug.Print "gld_rp_id debit =" & gld_rp_id
                        ws_gld_data.Cells(gld_rp_id,1)  = gld_rp_id 'id
                        ws_gld_data.Cells(gld_rp_id,2)  = gl_id 'general_ledger_id
                        ws_gld_data.Cells(gld_rp_id,3)  = coaCashAndBank 'chart_account_id Cash and Bank
                        ws_gld_data.Cells(gld_rp_id,4)  = 1 'company_id
                        ws_gld_data.Cells(gld_rp_id,5)  = branchId 'branch id
                        ws_gld_data.Cells(gld_rp_id,6)  = locationGroupId 'location group id
                        ws_gld_data.Cells(gld_rp_id,7)  = locationId 'location id
                        ws_gld_data.Cells(gld_rp_id,16) = "Receive Payment" 'type
                        ws_gld_data.Cells(gld_rp_id,17) = totalAmountPaid 'debit
                        ws_gld_data.Cells(gld_rp_id,18) = 0 'credit
                        ws_gld_data.Cells(gld_rp_id,19) = "Receive Payment # " & receiptNo 'memo
                        ws_gld_data.Cells(gld_rp_id,20) = customerId 'customer_id
                        ws_gld_data.Cells(gld_rp_id,24) = 1 'class_id
                        ws_gld_data.Cells(gld_rp_id,25) = 1 'is_active
                        ws_gld_data.Cells(gld_rp_id,26) = invoiceWeek 'week
                        ws_gld_data.Cells(gld_rp_id,27) = invoiceMonth 'month
                        ws_gld_data.Cells(gld_rp_id,28) = invoiceYear 'year
                        gld_rp_id = gld_rp_id + 1

                        ' debug.Print "gld_rp_id credit =" & gld_rp_id
                        ws_gld_data.Cells(gld_rp_id,1)  = gld_rp_id 'id
                        ws_gld_data.Cells(gld_rp_id,2)  = gl_id 'general_ledger_id
                        ws_gld_data.Cells(gld_rp_id,3)  = coaReceivable 'chart_account_id AR
                        ws_gld_data.Cells(gld_rp_id,4)  = 1 'company_id
                        ws_gld_data.Cells(gld_rp_id,5)  = branchId 'branch id
                        ws_gld_data.Cells(gld_rp_id,6)  = locationGroupId 'location group id
                        ws_gld_data.Cells(gld_rp_id,7)  = locationId 'location id
                        ws_gld_data.Cells(gld_rp_id,16) = "Receive Payment" 'type
                        ws_gld_data.Cells(gld_rp_id,17) = 0 'debit
                        ws_gld_data.Cells(gld_rp_id,18) = grandTotalPaid 'credit
                        ws_gld_data.Cells(gld_rp_id,19) = "Receive Payment # " & receiptNo 'memo
                        ws_gld_data.Cells(gld_rp_id,20) = customerId 'customer_id
                        ws_gld_data.Cells(gld_rp_id,24) = 1 'class_id
                        ws_gld_data.Cells(gld_rp_id,25) = 1 'is_active
                        ws_gld_data.Cells(gld_rp_id,26) = invoiceWeek 'week
                        ws_gld_data.Cells(gld_rp_id,27) = invoiceMonth 'month
                        ws_gld_data.Cells(gld_rp_id,28) = invoiceYear 'year
                    End If

                    If (totalDiscount > 0) Then
                        gld_rp_dis_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                        ' debug.Print "gld_rp_dis_id=" & gld_rp_dis_id
                        ws_gld_data.Cells(gld_rp_dis_id,1)  = gld_rp_dis_id 'id
                        ws_gld_data.Cells(gld_rp_dis_id,2)  = gl_id 'general_ledger_id
                        ws_gld_data.Cells(gld_rp_dis_id,3)  = coaPaymentDiscount 'chart_account_id application.VLOOKUP("Sales Income",chart_account_default_list2,3,FALSE)
                        ws_gld_data.Cells(gld_rp_dis_id,4)  = 1 'company_id
                        ws_gld_data.Cells(gld_rp_dis_id,5)  = branchId 'branch id
                        ws_gld_data.Cells(gld_rp_dis_id,6)  = locationGroupId 'location group id
                        ws_gld_data.Cells(gld_rp_dis_id,7)  = locationId 'location id
                        ws_gld_data.Cells(gld_rp_dis_id,16) = "Receive Payment" 'type
                        ws_gld_data.Cells(gld_rp_dis_id,17) = totalDiscount 'debit
                        ws_gld_data.Cells(gld_rp_dis_id,18) = 0 'credit
                        ws_gld_data.Cells(gld_rp_dis_id,19) = "Discount Receive Payment # " & receiptNo 'memo
                        ws_gld_data.Cells(gld_rp_dis_id,20) = customerId 'customer_id
                        ws_gld_data.Cells(gld_rp_dis_id,24) = 1 'class_id
                        ws_gld_data.Cells(gld_rp_dis_id,25) = 1 'is_active
                        ws_gld_data.Cells(gld_rp_dis_id,26) = invoiceWeek 'week
                        ws_gld_data.Cells(gld_rp_dis_id,27) = invoiceMonth 'month
                        ws_gld_data.Cells(gld_rp_dis_id,28) = invoiceYear 'year
                    End If
                    sirID = sirID + 1
                End If

                '*******update amount paid,discount , balance and status sales invoice*******
                LastRowSaleInvoiceDate = ws_inv.Cells(Rows.Count, "C").End(xlUp).Row 
                For j = 8 To LastRowSaleInvoiceDate
                    If ws_inv.Cells(j, 5).Value  = invoiceCode Then
                        ' debug.Print "Update Sale Invoice=> total dis=" & totalDiscount & " , totalAmountPaid=" & totalAmountPaid + totalAmountPaidBefore & " , balance=" & totalBalance & ",statusInvoice=" & statusInvoice
                        ws_inv.Cells(j,27) = ws_inv.Cells(j,27) + totalDiscount 'total discount
                        ws_inv.Cells(j,28) = ws_inv.Cells(j,28) + totalAmountPaid 'total amount paid
                        ws_inv.Cells(j,29) = totalBalance 'balance
                        ws_inv.Cells(j,30) = statusInvoice 'sale invoice status

                        Call setStatusIconReceivePayment(ws_inv.Cells(j,30)) 'set icon color base on invoice status type => -1 = Edit, 1  = issue, 2  = fullfilied ,3  = partial
                    End If
                Next j

                Call clearReceivePayment(activeWorkbook.Sheets("AddReceivePayment")) 'Reset field input after save
                ws_inv_receipt.Select
                ws_inv_receipt.Range("C" & ws_inv_receipt.Cells(Rows.Count, "C").End(xlUp).Row & ":" & "AD" & ws_inv_receipt.Cells(Rows.Count, "C").End(xlUp).Row).Select 'Select New insert invoice row
                wbDatabase.Save
                OnEnd
                MsgBox "Receive payment successful saved."
                Exit Sub
            Else
                MsgBox "Please input amount paid, try again."
                Exit Sub
            End If
        Else
            exit sub
        End If
    End If
End Sub

Private Function setStatusIconReceivePayment(rngInvoice As Range)
    Dim iset As IconSetCondition

    ' rngInvoice.Select
    rngInvoice.FormatConditions.Delete
    Set iset = rngInvoice.FormatConditions.AddIconSetCondition

    'Select the traffic lights iconset
    With iset
        .IconSet        = ActiveWorkbook.IconSets(xl4TrafficLights)
        .ReverseOrder   = False
        .ShowIconOnly   = True
    End With

    ' xlIconBlackCircleWithBorder , xlIconGrayCircle , xlIconGreenCircle , xlIconRedCircleWithBorder , xlIconPinkCircle , xlIconYellowCircle , xlIconGreenCheckSymbol , xlIconRedCrossSymbol , xlIconYellowExclamationSymbol , xlIconWhiteCircleAllWhiteQuarters
    ' Sales Invoice Status=> -1 = Edit, 1  = issue, 2  = fulfilled ,3  = partial

    With iset.IconCriteria(2)
        .Icon       = xlIconBlackCircleWithBorder
        .Type       = xlConditionValueFormula
        .Operator   = xlGreaterEqual
        .Value      = "=-1" 'edit
    End With

    With iset.IconCriteria(2)
        .Icon       = xlIconRedCircleWithBorder
        .Type       = xlConditionValueFormula
        .Operator   = xlGreaterEqual
        .Value      = "=1" 'issue
    End With

    With iset.IconCriteria(3)
        .Icon       = xlIconGreenCheckSymbol
        .Type       = xlConditionValueFormula
        .Operator   = xlGreaterEqual
        .Value      = "=2" 'fulfilled
    End With

    With iset.IconCriteria(4)
        .Icon       = xlIconYellowCircle
        .Type       = xlConditionValueFormula
        .Operator   = xlGreaterEqual
        .Value      = "=3" 'partial
    End With
End Function

Function clearReceivePayment(ws_rp As Worksheet)
    ws_rp.Range("rp_invoice_no").Value      = ""
    ws_rp.Range("rp_amount_paid_kh").Value  = ""
    ws_rp.Range("rp_amount_paid_en").Value  = ""
    ws_rp.Range("rp_discount_kh").Value     = ""
    ws_rp.Range("rp_discount_en").Value     = ""
    ws_rp.Range("rp_description").Value     = ""
End Function

Sub copyReceivePayment()
    Sheets("AddReceivePayment").Range("rp_print_receipt").Copy
End Sub

Sub CurrentDate()
    ActiveSheet.Range("rp_current_date").Value = Format(Now(), "mm/dd/yyyy")
End Sub

Sub ExchangeRateReceivePayment()
    ActiveSheet.Range("rp_exchange_rate").Value = Sheets("NBC_Exchange_Rate").Range("ExchangeRateToday").Value
End Sub

' Sub updateStatusColorReceivePayment()
'     Dim indexColor AS Integer
'     Dim status AS Integer
'     Dim rp AS Worksheet

'     Set rp = Sheets("ReceivePayment")
'     status = ws_rp.Range("rp_status").Value
'     If status = -1 Then
'         indexColor = 1
'     ElseIf status = 1 Then
'         indexColor = 3
'     ElseIf status = 2 Then
'         indexColor = 43
'     ElseIf status = 3 Then
'         indexColor = 44
'     End If

'     With ws_rp.Range("rp_status_note").Interior 'inv_status_note = Z7
'         .ColorIndex = indexColor
'     End With

'     With ws_rp.Range("rp_status_note").Font
'         .Bold = False
'         .ColorIndex = 2
'         .Name = "Roboto"
'         .Size = 12
'     End With
' End Sub

' Function updateBalanceSalesInvoice(obj)
'     Dim LastRow, lastCol, i As Long
'     Dim ws,rp As Worksheet
'     Dim invoiceNo AS Integer 
'     Dim totalAmountPaid,totalBalance,totalAmountDue,totalAmountPaidBefore AS Double
'     Dim statusInvoice AS Integer

'     Set ws = Sheets("SaleInvoiceData")
'     Set rp = Sheets("ReceivePayment")

'     invoiceNo       = ws_rp.Range("rp_invoice_no").Value
'     totalAmountDue  = ws_rp.Range("rp_amount_due").Value
'     totalAmountPaid = ws_rp.Range("rp_total_amount_paid").Value
'     totalBalance    = ws_rp.Range("rp_total_balance").Value
'     totalAmountPaidBefore = ws_rp.Range("rp_amount_paid_before").Value

'     'Sales Invoice Status=> -1 = Edit, 1  = issue, 2  = fullfilied ,3  = partial
'     If totalAmountPaid = totalAmountDue Then
'         statusInvoice = 2
'     ElseIf totalAmountPaid < totalAmountDue Then
'         statusInvoice = 3
'     Else
'         statusInvoice = 1
'     End If

'     LastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row 
'     lastCol = ws.Cells(6, Columns.Count).End(XlToLeft).column 'start from row 6 column B
 
'     For i = 6 To LastRow
'         If ws.Cells(i, 2).Value = invoiceNo*1 Then
'             'ws.Rows(i).EntireRow.copy
'             ws.Cells(i, 14).Value = totalAmountPaid + totalAmountPaidBefore
'             ws.Cells(i, 15).Value = totalBalance
'             ws.Cells(i, 17).Value = statusInvoice
'         End If
'     Next i
' End Function

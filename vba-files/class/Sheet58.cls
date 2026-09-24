VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub OnStart()
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual 'xlCalculationManual , xlAutomatic
    Application.DisplayStatusBar = False
    Application.ScreenUpdating = False
    ' Application.DisplayAlerts = False
    ' Application.StatusBar = False
    Application.AskToUpdateLinks = False
End Sub

Private Sub OnEnd()
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    ' Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ' Application.StatusBar = True
    Application.AskToUpdateLinks = True
End Sub

Private Sub InvoiceComboBox_Change()
    ' OnStart
    'workbook , worksheet variable
    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_file_setting,ws_print_invoice, ws_gld, ws_gl,ws, ws_inv_data,ws_inv_item_data As Worksheet

    'loop variable
    Dim i,j,k,m, LastRow As Integer

    'customer variable
    Dim lastRowCustomer As Integer
    Dim customerName, customerId, customerNameEn,customerNameKh,customerAddressKh,customerTel,customerEmail,customerTinNumber As String

    'invoice variable
    Dim invoiceDate,invoiceDueDate As Date
    Dim statusNote,priceTypeId,discountTypeId,invoiceCode,salesInvoiceNote As String
    Dim vatCal, vatPercent, invoiceId, netDays, paymentTermId, conversion, invId, invDetailId,exchangeRate,lastInvoiceDetailId As Integer

    'amount variable
    Dim totalAmount, unitPrice, totalPrice, totalDeposit,totalDiscount,grandTotalAmount,totalDisItem,disItem,totalVat As Double

    'Product variable
    Dim countInvoiceDetailItem,countInvoiceNo,qtyStock, productId, qty, qtyFree,convertion, lastRowUom As Integer
    Dim productCode, productName, productUoM As String
    Dim rngUom, rangeInvoiceList,rangeInvoiceDetail,rangePrintInvItem,rangeInvDetailId As Range

    'branch variablbe
    Dim branchListRange As Range
    Dim lastRowBranch As Integer
    Dim branchName, branchId, bankName,accountName,accountNo,branchAddressKh,branchTel,branchEmail As String

    'setting variable
    Dim customerListRange,paymentTermRange As Range
    Dim databasePath, databaseFile As String

    Set activeWorkbook     = Workbooks("index.xlsm")
    Set ws_print_invoice   = activeWorkbook.Sheets("PrintInvoice")
    Set rangePrintInvItem  = ws_print_invoice.Range("print_inv_items")

    If (Me.InvoiceComboBox.Value <> "Please select invoice code" And Me.InvoiceComboBox.Value <> "") Then 'checkDatabaseFile = TRUE And 
        rangePrintInvItem.ClearContents
        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath        = ws_file_setting.Range("database_path").Value
        databaseFile        = databasePath & "database_kscm.xlsm"

        ' If (IsWorkbookOpen(databaseFile) = FALSE) Then
        '     Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        ' End If

        Set wbDatabase          = Workbooks("database_kscm.xlsm")
        Set ws_inv_data         = wbDatabase.Sheets("SaleInvoiceData")
        Set ws_inv_item_data    = wbDatabase.Sheets("SalesByItemData")
        lastInvoiceDetailId     = ws_inv_item_data.Cells(Rows.Count, "C").End(xlUp).Row

        Set rangeInvoiceList    = ws_inv_data.Range("C8:AI" & lastInvoiceDetailId) 'invoice_list2
        Set rangeInvDetailId    = ws_inv_item_data.Range("C8:C" & lastInvoiceDetailId)
        Set rangeInvoiceDetail  = ws_inv_item_data.Range("C8:Z" & lastInvoiceDetailId)
        
        countInvoiceDetailItem  = Application.CountA(rangeInvDetailId)
        invoiceId               = CInt(Split(Me.InvoiceComboBox.Value,"#")(1))
        invoiceCode             = Split(Me.InvoiceComboBox.Value,"#")(0)

        customerId        = Application.VLookup(invoiceId, rangeInvoiceList, 12, False)
        invoiceDate       = Format(Application.VLookup(invoiceId, rangeInvoiceList, 4, False), "yyyy/mm/dd")
        paymentTermId     = Application.VLookup(invoiceId, rangeInvoiceList, 13, False)
        branchId          = Application.VLookup(invoiceId, rangeInvoiceList, 6, False)
        branchName        = ""
        customerName      = Application.VLookup(invoiceId, rangeInvoiceList, 17, False)
        salesInvoiceNote  = Application.VLookup(invoiceId, rangeInvoiceList, 29, False)

        vatPercent        = Application.VLookup(invoiceId, rangeInvoiceList, 21, False)
        vatCal            = Application.VLookup(invoiceId, rangeInvoiceList, 22, False)
        totalVat          = Application.VLookup(invoiceId, rangeInvoiceList, 24, False)

        exchangeRate      = Application.VLookup(invoiceId, rangeInvoiceList, 18, False)
        totalAmount       = Application.VLookup(invoiceId, rangeInvoiceList, 23, False) 
        totalDiscount     = Application.VLookup(invoiceId, rangeInvoiceList, 25, False)
        totalDeposit      = Application.VLookup(invoiceId, rangeInvoiceList, 26, False)
        grandTotalAmount  = Application.VLookup(invoiceId, rangeInvoiceList, 27, False)

        lastRowCustomer = activeWorkbook.Sheets("CustomerList").Cells(Rows.Count, "C").End(xlUp).Row
        lastRowBranch   = activeWorkbook.Sheets("Branch").Cells(Rows.Count, "C").End(xlUp).Row
        lastRowUom      = activeWorkbook.Sheets("UoM").Cells(Rows.Count, "C").End(xlUp).Row

        Set customerListRange   = activeWorkbook.Sheets("CustomerList").Range("C5:P" & lastRowCustomer)
        Set branchListRange     = activeWorkbook.Sheets("Branch").Range("C7:AF" & lastRowBranch)
        Set paymentTermRange    = activeWorkbook.Sheets("Setting").Range("payment_term_list")
        Set rngUom              = activeWorkbook.Sheets("UoM").Range("C8:I" & lastRowUom)
        
        netDays                 = Application.VLookup(paymentTermId, paymentTermRange, 3, False)
        invoiceDueDate          = 0
        If netDays <> "" Then
            invoiceDueDate      = DateAdd("d", netDays, invoiceDate)
        End If

        customerNameEn          = Application.VLookup(customerId, customerListRange, 5, False)
        customerNameKh          = Application.VLookup(customerId, customerListRange, 6, False)
        customerAddressKh       = Application.VLookup(customerId, customerListRange, 9, False)
        customerTel             = Application.VLookup(customerId, customerListRange, 7, False)
        customerEmail           = Application.VLookup(customerId, customerListRange, 10, False)
        customerTinNumber       = Application.VLookup(customerId, customerListRange, 12, False)

        bankName                = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(0)
        accountName             = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(1)
        accountNo               = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(2)
        branchAddressKh         = Application.VLookup(branchId, branchListRange, 22, False)
        branchTel               = Application.VLookup(branchId, branchListRange, 18, False)
        branchEmail             = Application.VLookup(branchId, branchListRange, 20, False)

        'Invoice Information
        ws_print_invoice.Range("print_inv_code").Value = invoiceCode
        ws_print_invoice.Range("print_inv_date").Value = invoiceDate
        ws_print_invoice.Range("print_inv_due_date").Value = invoiceDueDate
        ws_print_invoice.Range("print_inv_note").Value = salesInvoiceNote
        ws_print_invoice.Range("print_inv_exchange_rate").Value = exchangeRate
        ws_print_invoice.Range("print_inv_payment_term_id").Value = paymentTermId
        ws_print_invoice.Range("print_inv_payment_net_days").Value = netDays
    
        'Amount Information
        ws_print_invoice.Range("print_inv_sub_total").Value = totalAmount
        ws_print_invoice.Range("print_inv_sub_total_kh").Value = totalAmount * exchangeRate
        ws_print_invoice.Range("print_inv_toal_discount").Value = totalDiscount
        ws_print_invoice.Range("print_inv_toal_discount_kh").Value = totalDiscount * exchangeRate
        ws_print_invoice.Range("print_inv_toal_deposit").Value = totalDeposit
        ws_print_invoice.Range("print_inv_toal_vat").Value = totalVat
        ws_print_invoice.Range("print_inv_grand_total").Value = grandTotalAmount
        ws_print_invoice.Range("print_inv_grand_total_kh").Value = grandTotalAmount * exchangeRate

        'Branch information
        ws_print_invoice.Range("print_inv_bank_name").Value = bankName
        ws_print_invoice.Range("print_inv_account_name").Value = accountName
        ws_print_invoice.Range("print_inv_account_no").Value = accountNo
        ws_print_invoice.Range("print_inv_branch_address_kh").Value = branchAddressKh
        ws_print_invoice.Range("print_inv_branch_telephone").Value = branchTel
        ws_print_invoice.Range("print_inv_branch_email").Value = branchEmail
        
        'Customer Information
        ws_print_invoice.Range("print_inv_customer_name_kh").Value = customerNameKh
        ws_print_invoice.Range("print_inv_customer_name_en").Value = customerNameEn
        ws_print_invoice.Range("print_inv_customer_address_kh").Value = customerAddressKh
        ws_print_invoice.Range("print_inv_customer_tel").Value = customerTel
        ws_print_invoice.Range("print_inv_customer_email").Value = customerEmail
        ws_print_invoice.Range("print_inv_customer_tin_number").Value = customerTinNumber
        
        ws_print_invoice.Range("C27" & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = False 'Show all row in range print invoice
        
        If invoiceId > 0 AND countInvoiceDetailItem > 0 Then
            k = 0
            m = 0
            ' debug.Print "countInvoiceDetailItem=" & countInvoiceDetailItem
            For j = 1 To countInvoiceDetailItem
                If (rangeInvoiceDetail.Cells(j, 2) = invoiceId) Then
                    ' disItem             = rangeInvoiceDetail.Cells(j,6)
                    ' totalDisItem        = rangeInvoiceDetail.Cells(j,6)
                    productCode           = rangeInvoiceDetail.Cells(j, 8)
                    productName           = rangeInvoiceDetail.Cells(j, 9) 'MID(Split(rangeInvoiceDetail.Cells(j, 7),"#")(0), 7, 100)
                    qty                   = rangeInvoiceDetail.Cells(j, 10)
                    qtyFree               = rangeInvoiceDetail.Cells(j, 11)
                    productUoM            = Application.VLookup(MID(rangeInvoiceDetail.Cells(j, 12), 1, 3) *1, rngUom,5, False)
                    unitPrice             = rangeInvoiceDetail.Cells(j, 18)
                    totalPrice            = rangeInvoiceDetail.Cells(j, 19)
                 
                    Dim criteriaRange1, criteriaRange2 As Range
                    Dim countItem As Integer
                    Set criteriaRange1 = ws_inv_item_data.Range("D8:D" & countInvoiceDetailItem + 7) 'column invoice id
                    Set criteriaRange2 = ws_inv_item_data.Range("AB8:AB" & countInvoiceDetailItem + 7) 'column status
                    countItem = Application.COUNTIFS(criteriaRange1,invoiceId,criteriaRange2,1)
                    ' debug.Print "countItem=" & countItem
                    If (qty > 0) Then
                        k = k + 1
                        If (countItem >= 26 AND countItem <= 104) Then '26, 52, 78, 104
                            If (k Mod 26 = 0) Then 
                                ' Debug.Print "***************k=" & k
                                k = k + 1
                            End If
                        Else
                            If (k Mod 13 = 0) Then ' k = 13 or k = 26 or k = 39 or k = 52 or k = 65
                                ' Debug.Print "***************k=" & k
                                k = k + 14
                            End If
                        End If
                        ' Debug.Print "k=" & k
                        rangePrintInvItem.Cells(k,1) = productCode
                        rangePrintInvItem.Cells(k,2) = productName
                        rangePrintInvItem.Cells(k,5) = productUoM
                        rangePrintInvItem.Cells(k,6) = qty
                        rangePrintInvItem.Cells(k,7) = unitPrice
                        rangePrintInvItem.Cells(k,8) = totalPrice
                        rangePrintInvItem.Cells(k,9) = totalPrice * exchangeRate
                    End If
                End If
            Next j

            lastRow = k + 27
            ws_print_invoice.Range("C" & lastRow & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
                
            ' OnEnd
        End If
    Else
        ' MsgBox "Database file doesn't exist."
        ' Exit Sub
        rangePrintInvItem.ClearContents
    End If

End Sub

Private Sub TaxInvoiceComboBox_Change()
    ' OnStart
    'workbook, worksheet variable
    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_file_setting, ws_print_invoice, ws_cus, ws_gld, ws_gl, ws, ws_inv_data, ws_inv_item_data As Worksheet

    'loop variable
    Dim i,j,k,n,m As Integer

    'customer variable
    Dim customerListRange As Range
    Dim lastCustomerId,lastRowCustomer As Integer
    Dim customerName As String
    Dim customerNameEn,customerNameKh,customerAddressKh,customerTel,customerEmail,customerTinNumber As String

    'invoice variable
    Dim invoiceDate,invoiceDueDate As Date
    Dim statusNote, invoiceCode, salesInvoiceNote As String
    Dim rangeInvoiceList,rangeInvoiceDetail,paymentTermRange, rangePrintInvItem, rangeInvDetailId, rngInvoiceList As Range
    Dim lastInvoiceDetailId, lastInvoiceId, invoiceId, netDays, paymentTermId, exchangeRate As Integer

    'amount variable
    Dim totalDis, totalVat, subTotal, unitPrice, totalPrice, totalDeposit,grandTotalAmount As Double

    'Product variable
    Dim productCode, productName, productUoM As String
    Dim rngUom, item_row As Range
    Dim conversion, countInvoiceDetailItem, qty, qtyFree, lastRowUom As Integer

    'branch variablbe
    Dim bankName,accountName,accountNo,branchAddressKh,branchTel,branchEmail,branchName As String
    Dim branchListRange As Range
    Dim branchId,lastRowRngPrintInv,lastRow, lastRowBranch As Integer
    
    'setting variable
    Dim priceTypeId,discountTypeId, databasePath, databaseFile As String
    
    Set activeWorkbook      = Workbooks("index.xlsm")
    Set ws_print_invoice    = activeWorkbook.Sheets("PrintInvoice")
    Set rangePrintInvItem   = ws_print_invoice.Range("print_inv_items")

    lastRowCustomer = activeWorkbook.Sheets("CustomerList").Cells(Rows.Count, "C").End(xlUp).Row
    lastRowBranch   = activeWorkbook.Sheets("Branch").Cells(Rows.Count, "C").End(xlUp).Row
    lastRowUom      = activeWorkbook.Sheets("UoM").Cells(Rows.Count, "C").End(xlUp).Row

    Set customerListRange   = activeWorkbook.Sheets("CustomerList").Range("C5:P" & lastRowCustomer)
    Set branchListRange     = activeWorkbook.Sheets("Branch").Range("C7:AF" & lastRowBranch)
    Set paymentTermRange    = activeWorkbook.Sheets("Setting").Range("payment_term_list")
    Set rngUom              = activeWorkbook.Sheets("UoM").Range("C8:I" & lastRowUom)
    
    If (Me.TaxInvoiceComboBox.Value <> "Please select tax invoice code" And Me.TaxInvoiceComboBox.Value <> "") Then
        rangePrintInvItem.ClearContents

        ws_print_invoice.Range("C27" & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = False 'Show all row in range print invoice

        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath        = ws_file_setting.Range("database_path").Value
        databaseFile        = databasePath & "database_kscm.xlsm"

        ' If (IsWorkbookOpen(databaseFile) = FALSE) Then
        '     Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        ' End If

        Set wbDatabase          = Workbooks("database_kscm.xlsm")
        Set ws_inv_data         = wbDatabase.Sheets("SaleInvoiceData")
        Set ws_inv_item_data    = wbDatabase.Sheets("SalesByItemData")
        lastInvoiceId           = ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row
        lastInvoiceDetailId     = ws_inv_item_data.Cells(Rows.Count, "C").End(xlUp).Row

        Set rangeInvoiceList    = ws_inv_data.Range("D8:AI" & lastInvoiceDetailId) 'invoice_list2
        Set rangeInvDetailId    = ws_inv_item_data.Range("C8:C" & lastInvoiceDetailId)
        Set rangeInvoiceDetail  = ws_inv_item_data.Range("C8:Z" & lastInvoiceDetailId)

        countInvoiceDetailItem  = Application.CountA(rangeInvDetailId)
        invoiceCode             = Me.TaxInvoiceComboBox.Value

        Set rngInvoiceList      = ws_inv_data.Range("C8:AI" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row)
        Set ws_cus              = activeWorkbook.Sheets("CustomerList")
        lastCustomerId          = ws_cus.Cells(Rows.Count, "C").End(xlUp).Row
        Set customerListRange   = ws_cus.Range("C5:P" & lastCustomerId)

        Dim arrCustomerId As Variant
        Dim arrCusName, arrInvoiceCode As String

        arrCusName     = ""
        arrInvoiceCode = ""
        exchangeRate   = Application.VLookup(invoiceCode, rangeInvoiceList, 17, False)
        Dim customerId As Variant
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        For i = 1 To lastInvoiceId
            If rngInvoiceList.Cells(i, 1) <> "" Then
                If rngInvoiceList.Cells(i, 2) = invoiceCode Then
                    'Get unique invoice code value
                    If Not dict.Exists(rngInvoiceList.Cells(i, 3)) And rngInvoiceList.Cells(i, 3) <> "" Then
                        arrInvoiceCode = arrInvoiceCode & rngInvoiceList.Cells(i, 3) & " , "
                        dict.Add rngInvoiceList.Cells(i, 3), True ' Track that this value is processed
                    End If
                    invoiceDate = rngInvoiceList.Cells(i, 4)

                    'Count number of customer
                    Dim invoiceCodeCriteriaRng As Range
                    Dim countTaxCode As Integer
                    Set invoiceCodeCriteriaRng = ws_inv_data.Range("D8:D" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row) 'column invoice code
                    countTaxCode = Application.COUNTIFS(invoiceCodeCriteriaRng,invoiceCode)

                    'Get unique customer name value 
                    customerId = rngInvoiceList.Cells(i, 12) 
                    If (countTaxCode = 1) Then 
                        arrCusName = Application.VLookup(customerId, customerListRange, 6, False) & " , "
                    Else
                        ' If CInt(rngInvoiceList.Cells(i, 12)) <> CInt(rngInvoiceList.Cells(i + 1, 12)) Then
                        If Not dict.Exists(customerId) And customerId <> "" Then
                            arrCusName = arrCusName & Application.VLookup(customerId, customerListRange, 6, False) & " , "
                            dict.Add customerId, True ' Track that this value is processed
                        End If
                        ' End If
                    End If
                End If
            End If
        Next i
        If Len(arrCusName) > 0 Then
            arrCusName = Left(arrCusName, Len(arrCusName) - 2)
        End If
        If Len(arrInvoiceCode) > 0 Then
            arrInvoiceCode = Left(arrInvoiceCode, Len(arrInvoiceCode) - 2)
        End If

        'Branch information
        branchId = 1
        bankName                = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(0)
        accountName             = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(1)
        accountNo               = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(2)
        branchAddressKh         = Application.VLookup(branchId, branchListRange, 22, False)
        branchTel               = Application.VLookup(branchId, branchListRange, 18, False)
        branchEmail             = Application.VLookup(branchId, branchListRange, 20, False)

        ws_print_invoice.Range("print_inv_bank_name").Value          = bankName
        ws_print_invoice.Range("print_inv_account_name").Value       = accountName
        ws_print_invoice.Range("print_inv_account_no").Value         = accountNo
        ws_print_invoice.Range("print_inv_branch_address_kh").Value  = branchAddressKh
        ws_print_invoice.Range("print_inv_branch_telephone").Value   = branchTel
        ws_print_invoice.Range("print_inv_branch_email").Value       = branchEmail

        'Add customer information
        ws_print_invoice.Range("print_inv_customer_name_kh").Value    = arrCusName
        ws_print_invoice.Range("print_inv_customer_name_en").Value    = ""
        ws_print_invoice.Range("print_inv_customer_address_kh").Value = ""
        ws_print_invoice.Range("print_inv_customer_tel").Value        = ""
        ws_print_invoice.Range("print_inv_customer_email").Value      = ""
        ws_print_invoice.Range("print_inv_customer_tin_number").Value = ""

        'Add invoice information
        ws_print_invoice.Range("print_inv_code").Value             = invoiceCode
        ws_print_invoice.Range("print_inv_date").Value             = invoiceDate
        ws_print_invoice.Range("print_inv_due_date").Value         = ""
        ws_print_invoice.Range("print_inv_note").Value             = arrInvoiceCode
        ws_print_invoice.Range("print_inv_exchange_rate").Value    = exchangeRate
        ws_print_invoice.Range("print_inv_payment_term_id").Value  = ""
        ws_print_invoice.Range("print_inv_payment_net_days").Value = 0

        totalDis     = Application.SUMIFS(ws_inv_data.Range("AA8:AA" & lastInvoiceId),ws_inv_data.Range("D8:D" & lastInvoiceId),invoiceCode,ws_inv_data.Range("AD8:AD" & lastInvoiceId),">0") 'SUMIFS(AA8:AA11,D8:D11,D8,AD8:AD11,">0")
        totalVat     = Application.SUMIFS(ws_inv_data.Range("Z8:Z" & lastInvoiceId),ws_inv_data.Range("D8:D" & lastInvoiceId),invoiceCode,ws_inv_data.Range("AD8:AD" & lastInvoiceId),">0")  'SUMIFS(Z8:Z11,D8:D11,D8,AD8:AD11,">0")
        totalDeposit = 0

        If countInvoiceDetailItem > 0 Then
            k = 0
            subTotal = 0
            For j = 1 To countInvoiceDetailItem
                If (rangeInvoiceDetail.Cells(j, 4) = invoiceCode) Then
                    productCode           = rangeInvoiceDetail.Cells(j, 8)
                    productName           = rangeInvoiceDetail.Cells(j, 9) 'MID(Split(rangeInvoiceDetail.Cells(j, 7),"#")(0), 7, 100)
                    qty                   = rangeInvoiceDetail.Cells(j, 10)
                    qtyFree               = rangeInvoiceDetail.Cells(j, 11)
                    productUoM            = Application.VLookup(MID(rangeInvoiceDetail.Cells(j, 12), 1, 3) *1, rngUom,5, False)
                    unitPrice             = rangeInvoiceDetail.Cells(j, 18)
                    totalPrice            = rangeInvoiceDetail.Cells(j, 19)

                    Dim criteriaRange1, criteriaRange2 As Range
                    Dim countItem As Integer
                    Set criteriaRange1 = ws_inv_item_data.Range("F8:F" & countInvoiceDetailItem + 7) 'column invoice id
                    Set criteriaRange2 = ws_inv_item_data.Range("AB8:AB" & countInvoiceDetailItem + 7) 'column status
                    countItem = Application.COUNTIFS(criteriaRange1,invoiceCode,criteriaRange2,1)

                    If (qty > 0) Then
                        k = k + 1
                        If (countItem >= 26 AND countItem <= 104) Then '26, 52, 78, 104
                            If (k Mod 26 = 0) Then 
                                ' Debug.Print "***************k=" & k
                                k = k + 1
                            End If
                        Else
                            If (k Mod 13 = 0) Then ' k = 13
                                ' Debug.Print "***************k=" & k
                                k = k + 14
                            End If
                        End If

                        ' Debug.Print "=>k=" & k
                        rangePrintInvItem.Cells(k,1) = productCode
                        rangePrintInvItem.Cells(k,2) = productName
                        rangePrintInvItem.Cells(k,5) = productUoM
                        rangePrintInvItem.Cells(k,6) = qty
                        rangePrintInvItem.Cells(k,7) = unitPrice
                        rangePrintInvItem.Cells(k,8) = totalPrice
                        rangePrintInvItem.Cells(k,9) = totalPrice * exchangeRate
                        subTotal = subTotal + totalPrice
                    End If
                End If
            Next j

            lastRow = k + 27
            ws_print_invoice.Range("C" & lastRow & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row

            ws_print_invoice.Range("print_inv_sub_total").Value        = subTotal
            ws_print_invoice.Range("print_inv_sub_total_kh").Value     = subTotal * exchangeRate
            ws_print_invoice.Range("print_inv_toal_discount").Value    = totalDis
            ws_print_invoice.Range("print_inv_toal_discount_kh").Value = totalDis * exchangeRate
            ws_print_invoice.Range("print_inv_toal_deposit").Value     = totalDeposit
            ws_print_invoice.Range("print_inv_toal_vat").Value         = totalVat

            grandTotalAmount = Format(subTotal + totalVat - totalDeposit - totalDis, "0.00")

            ws_print_invoice.Range("print_inv_grand_total").Value = grandTotalAmount
            ws_print_invoice.Range("print_inv_grand_total_kh").Value = grandTotalAmount * exchangeRate
        End If
    Else
        'MsgBox "Database file doesn't exist."
        'Exit Sub
        rangePrintInvItem.ClearContents
    End If
    ' OnEnd
End Sub

Private Sub Worksheet_Activate()
    ' OnStart
    ' ' Invoice Information
    ' ws_print_invoice.Range("print_inv_code").Value = ""
    ' ws_print_invoice.Range("print_inv_date").Value = ""
    ' ws_print_invoice.Range("print_inv_due_date").Value = ""
    ' ws_print_invoice.Range("print_inv_note").Value = ""
    ' ws_print_invoice.Range("print_inv_exchange_rate").Value = ""
    ' ws_print_invoice.Range("print_inv_company_name_kh").Value = ""
    ' ws_print_invoice.Range("print_inv_company_name_en").Value = ""
    ' ws_print_invoice.Range("print_inv_company_tin").Value = ""
    ' ws_print_invoice.Range("print_inv_payment_term_id").Value = ""
    ' ws_print_invoice.Range("print_inv_payment_net_days").Value = ""

    ' ws_print_invoice.Range("print_inv_sub_total").Value = 0
    ' ws_print_invoice.Range("print_inv_sub_total_kh").Value = 0
    ' ws_print_invoice.Range("print_inv_toal_discount").Value = 0
    ' ws_print_invoice.Range("print_inv_toal_discount_kh").Value = 0
    ' ws_print_invoice.Range("print_inv_toal_deposit").Value = 0
    ' ws_print_invoice.Range("print_inv_toal_vat").Value = 0
    ' ws_print_invoice.Range("print_inv_grand_total").Value = 0
    ' ws_print_invoice.Range("print_inv_grand_total_kh").Value = 0

    ' ' Branch information bankName,accountName,accountNo,branchAddressKh,branchTel,branchEmail
    ' ws_print_invoice.Range("print_inv_bank_name").Value = ""
    ' ws_print_invoice.Range("print_inv_account_name").Value = ""
    ' ws_print_invoice.Range("print_inv_account_no").Value = ""
    ' ws_print_invoice.Range("print_inv_branch_address_kh").Value = ""
    ' ws_print_invoice.Range("print_inv_branch_telephone").Value = ""
    ' ws_print_invoice.Range("print_inv_branch_email").Value = ""

    ' ' ' Customer Information
    ' ws_print_invoice.Range("print_inv_customer_name_kh").Value = ""
    ' ws_print_invoice.Range("print_inv_customer_name_en").Value = ""
    ' ws_print_invoice.Range("print_inv_customer_address_kh").Value = ""
    ' ws_print_invoice.Range("print_inv_customer_tel").Value = ""
    ' ws_print_invoice.Range("print_inv_customer_email").Value = ""
    ' ws_print_invoice.Range("print_inv_customer_tin_number").Value = ""

    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_print_invoice, ws_inv_data,ws_file_setting As Worksheet
    Dim rowInvoiceCode, rowInvoiceId, rowTaxInvoiceCode As Range
    Dim i,k, CountaResultCell,lastInvoiceId As Integer
    Dim databasePath, databaseFile As String

    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_print_invoice = activeWorkbook.Sheets("PrintInvoice")
    
    ws_print_invoice.TaxInvoiceComboBox.Clear
    ws_print_invoice.InvoiceComboBox.Clear
    ws_print_invoice.InvoiceComboBox.AddItem "Please select invoice code"
    ws_print_invoice.TaxInvoiceComboBox.AddItem "Please select tax invoice code"

    If (checkDatabaseFile = TRUE) Then 
        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath        = ws_file_setting.Range("database_path").Value
        databaseFile        = databasePath & "database_kscm.xlsm"

        If (IsWorkbookOpen(databaseFile) = FALSE) Then
            ' Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        End If
        
        Set wbDatabase     = Workbooks("database_kscm.xlsm")
        Set ws_inv_data    = wbDatabase.Sheets("SaleInvoiceData")
        lastInvoiceId      = ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row - 7

        Set rowTaxInvoiceCode  = ws_inv_data.Range("D8:D" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row) 
        Set rowInvoiceCode     = ws_inv_data.Range("E8:E" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row)
        Set rowInvoiceId       = ws_inv_data.Range("C8:C" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row)

        For i = 1 To lastInvoiceId
            If rowInvoiceCode.Cells(i) <> "" Then
                ws_print_invoice.InvoiceComboBox.AddItem rowInvoiceCode(i) & " #" & rowInvoiceId(i)
                If rowTaxInvoiceCode(i) <> rowTaxInvoiceCode(i+1) Then
                    ws_print_invoice.TaxInvoiceComboBox.AddItem rowTaxInvoiceCode(i)
                End If
            End If
        Next i
    Else
        ' MsgBox "Database file doesn't exist."
        ' Exit Sub
    End If
    
    ' OnEnd
End Sub

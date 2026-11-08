VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet67"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'******sales by service data
' C.Id
' D.Invoice ID
' E.Tax Inv No
' F.Invoice No
' G.Invoice Date
' H.Customer ID
' I.Service code
' J.Service Description
' K.Qty
' L.Qty Free
' M.Uom Id
' N.Conversion
' O.Dis ID
' P.Dis Percent
' Q.Dis Amount
' R.Unit Cost
' S.Unit Price
' T.Total Price
' U.Remark
' V.Week
' W.Month
' X.Year
' Y.Status

' Private Sub OnStart()
'     Application.EnableEvents = False
'     Application.Calculation = xlCalculationManual 'xlCalculationManual , xlAutomatic
'     Application.DisplayStatusBar = False
'     Application.ScreenUpdating = False
'     ' Application.DisplayAlerts = False
'     ' Application.StatusBar = False
'     Application.AskToUpdateLinks = False
' End Sub

' Private Sub OnEnd()
'     Application.EnableEvents = True
'     Application.Calculation = xlCalculationAutomatic
'     Application.DisplayStatusBar = True
'     ' Application.DisplayAlerts = True
'     Application.ScreenUpdating = True
'     ' Application.StatusBar = True
'     Application.AskToUpdateLinks = True
' End Sub

' Private Sub InvoiceNormalComboBox_Change()
'     ' OnStart
'     'workbook , worksheet variable
'     Dim activeWorkbook, wbDatabase As Workbook
'     Dim ws_file_setting, ws_print_invoice, ws_gld, ws_gl, ws, ws_inv_data, ws_inv_item_data, ws_inv_service_data As Worksheet

'     'loop variable
'     Dim i, j, k, m, lastRow As Integer

'     'customer variable
'     Dim lastRowCustomer As Integer
'     Dim customerSocialNetwork, customerName, customerId, customerNameEn, customerNameKh, customerAddressKh, customerTel, customerEmail, customerTinNumber As String

'     'invoice variable
'     Dim invoiceDate, invoiceDueDate As Date
'     Dim statusNote, priceTypeId, discountTypeId, invoiceCode, salesInvoiceNote As String
'     Dim vatCal, vatPercent, invoiceId, netDays, paymentTermId, conversion, invId, invDetailId, exchangeRate, lastInvoiceDetailId, lastInvServiceId As Integer

'     'invoice detail variable
'     Dim rngInvoiceId, criteriaRange1, criteriaRange2 As Range
'     Dim countService, countItem As Integer

'     'amount variable
'     Dim totalAmount, unitPrice, totalPrice, totalDeposit, totalDiscount, grandTotalAmount, totalDisItem, disItem, totalVat As Double

'     'Product variable
'     Dim countInvoiceServiceItem, countInvoiceDetailItem, countInvoiceNo, qtyStock, productId, qty, qtyFree, convertion, lastRowUom As Integer
'     Dim productCode, productName, productUoM As String
'     Dim rngUom, rangeInvoiceList, rangeInvoiceDetail, rangeInvoiceService, rangePrintInvItem, rangeInvDetailId, rangeInvServiceId As Range

'     'branch variablbe
'     Dim branchListRange As Range
'     Dim lastRowBranch As Integer
'     Dim branchName, branchId, bankName, accountName, accountNo, branchAddressKh, branchTel, branchEmail As String

'     'setting variable
'     Dim customerListRange, paymentTermRange As Range
'     Dim databasePath, databaseFile As String

'     Set activeWorkbook = Workbooks("index.xlsm")
'     Set ws_print_invoice = activeWorkbook.Sheets("PrintInvoiceNormal")
'     Set rangePrintInvItem = ws_print_invoice.Range("print_inv_normal_items")
'     ' Set rngInvoiceId = ws_print_invoice.Range("A27:B117")

'     If (Me.InvoiceNormalComboBox.Value <> "Please select invoice code" And Me.InvoiceNormalComboBox.Value <> "") Then 'checkDatabaseFile = TRUE And
'         rangePrintInvItem.Value = "" ' rangePrintInvItem.ClearContents
'         Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
'         databasePath = ws_file_setting.Range("database_path").Value
'         databaseFile = databasePath & "database_kscm.xlsm"

'         ' If (IsWorkbookOpen(databaseFile) = FALSE) Then
'         '     Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
'         ' End If

'         Set wbDatabase = Workbooks("database_kscm.xlsm")
'         Set ws_inv_data = wbDatabase.Sheets("SaleInvoiceData")
'         Set ws_inv_item_data = wbDatabase.Sheets("SalesByItemData")
'         Set ws_inv_service_data = wbDatabase.Sheets("SalesByServiceData")

'         lastInvoiceDetailId = ws_inv_item_data.Cells(Rows.Count, "C").end(xlUp).row
'         lastInvServiceId = ws_inv_service_data.Cells(Rows.Count, "C").end(xlUp).row

'         Set rangeInvoiceList = ws_inv_data.Range("C8:AI" & lastInvoiceDetailId)    'invoice_list2

'         Set rangeInvDetailId = ws_inv_item_data.Range("C8:C" & lastInvoiceDetailId)
'         Set rangeInvoiceDetail = ws_inv_item_data.Range("C8:AB" & lastInvoiceDetailId)

'         Set rangeInvServiceId = ws_inv_service_data.Range("C8:C" & lastInvServiceId)
'         Set rangeInvoiceService = ws_inv_service_data.Range("C8:Y" & lastInvServiceId)

'         countInvoiceDetailItem = Application.CountA(rangeInvDetailId)
'         countInvoiceServiceItem = Application.CountA(rangeInvServiceId)
'         invoiceId = CInt(Split(Me.InvoiceNormalComboBox.Value, "#")(1))
'         invoiceCode = Split(Me.InvoiceNormalComboBox.Value, "#")(0)
'         paymentTermId = Application.VLookup(invoiceId, rangeInvoiceList, 13, False)

'         customerId        = Application.VLookup(invoiceId, rangeInvoiceList, 12, False)
'         invoiceDate = Format(Application.VLookup(invoiceId, rangeInvoiceList, 4, False), "yyyy/mm/dd")
        
'         ' branchId          = Application.VLookup(invoiceId, rangeInvoiceList, 6, False)
'         ' branchName        = ""
'         customerName      = Application.VLookup(invoiceId, rangeInvoiceList, 17, False)
'         salesInvoiceNote  = Application.VLookup(invoiceId, rangeInvoiceList, 29, False)

'         vatPercent = Application.VLookup(invoiceId, rangeInvoiceList, 21, False)
'         vatCal = Application.VLookup(invoiceId, rangeInvoiceList, 22, False)
'         totalVat = Application.VLookup(invoiceId, rangeInvoiceList, 24, False)

'         exchangeRate = Application.VLookup(invoiceId, rangeInvoiceList, 18, False)
'         totalAmount = Application.VLookup(invoiceId, rangeInvoiceList, 23, False)
'         totalDiscount = Application.VLookup(invoiceId, rangeInvoiceList, 25, False)
'         totalDeposit = Application.VLookup(invoiceId, rangeInvoiceList, 26, False)
'         grandTotalAmount = Application.VLookup(invoiceId, rangeInvoiceList, 27, False)

'         lastRowCustomer = activeWorkbook.Sheets("CustomerList").Cells(Rows.Count, "C").End(xlUp).Row
'         ' lastRowBranch   = activeWorkbook.Sheets("Branch").Cells(Rows.Count, "C").End(xlUp).Row
'         lastRowUom = activeWorkbook.Sheets("UoM").Cells(Rows.Count, "C").end(xlUp).row

'         Set customerListRange   = activeWorkbook.Sheets("CustomerList").Range("C5:P" & lastRowCustomer)
'         ' Set branchListRange     = activeWorkbook.Sheets("Branch").Range("C7:AF" & lastRowBranch)
'         Set paymentTermRange = activeWorkbook.Sheets("Setting").Range("payment_term_list")
'         Set rngUom = activeWorkbook.Sheets("UoM").Range("C8:I" & lastRowUom)
        
'         netDays = Application.VLookup(paymentTermId, paymentTermRange, 3, False)
'         invoiceDueDate = 0
'         If netDays <> "" Then
'             invoiceDueDate = DateAdd("d", netDays, invoiceDate)
'         End If

'         'customerNameEn         = CStr(Application.VLookup(customerId, customerListRange, 5, False))
'         customerNameKh          = CStr(Application.VLookup(customerId, customerListRange, 6, False))
'         customerTel             = CStr(Application.VLookup(customerId, customerListRange, 7, False))
'         customerSocialNetwork   = CStr(Application.VLookup(customerId, customerListRange, 8, False))
'         customerAddressKh       = CStr(Application.VLookup(customerId, customerListRange, 9, False))
'         customerEmail           = CStr(Application.VLookup(customerId, customerListRange, 10, False))
'         'customerTinNumber      = CStr(Application.VLookup(customerId, customerListRange, 12, False))

'         ' bankName                = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(0)
'         ' accountName             = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(1)
'         ' accountNo               = Split(Application.VLookup(branchId, branchListRange, 6, False),",")(2)
'         ' branchAddressKh         = Application.VLookup(branchId, branchListRange, 22, False)
'         ' branchTel               = Application.VLookup(branchId, branchListRange, 18, False)
'         ' branchEmail             = Application.VLookup(branchId, branchListRange, 20, False)

'         ' 'Invoice Information
'         ws_print_invoice.Range("print_inv_normal_id").Value = invoiceId
'         ws_print_invoice.Range("print_inv_normal_code").Value = invoiceCode
'         ws_print_invoice.Range("print_inv_normal_date").Value = invoiceDate
'         ws_print_invoice.Range("print_inv_normal_due_date").Value = invoiceDueDate
        
'         ws_print_invoice.Range("print_inv_normal_exchange_rate").Value = exchangeRate
'         ws_print_invoice.Range("print_inv_normal_payment_term_id").Value = paymentTermId
'         ws_print_invoice.Range("print_inv_normal_payment_net_days").Value = netDays
'         ws_print_invoice.Range("print_inv_normal_note").Value = salesInvoiceNote

'         ' 'Amount Information
'         ws_print_invoice.Range("print_inv_normal_sub_total").Value = totalAmount
'         ws_print_invoice.Range("print_inv_normal_total_discount").Value = totalDiscount
'         ws_print_invoice.Range("print_inv_normal_total_deposit").Value = totalDeposit
'         ws_print_invoice.Range("print_inv_normal_total_vat").Value = totalVat
'         ws_print_invoice.Range("print_inv_normal_grand_total").Value = grandTotalAmount
'         ws_print_invoice.Range("print_inv_normal_grand_total_kh").Value = grandTotalAmount * exchangeRate

'         ' 'Branch information
'         ' ws_print_invoice.Range("print_inv_normal_bank_name").Value = bankName
'         ' ws_print_invoice.Range("print_inv_normal_account_name").Value = accountName
'         ' ws_print_invoice.Range("print_inv_normal_account_no").Value = accountNo
'         ' ws_print_invoice.Range("print_inv_normal_branch_address_kh").Value = branchAddressKh
'         ' ws_print_invoice.Range("print_inv_normal_branch_telephone").Value = branchTel
'         ' ws_print_invoice.Range("print_inv_normal_branch_email").Value = branchEmail
        
'         ' 'Customer Information
'         ws_print_invoice.Range("print_inv_normal_customer_name_kh").Value = customerNameKh
'         ' ws_print_invoice.Range("print_inv_normal_customer_name_en").Value = customerNameEn
'         ws_print_invoice.Range("print_inv_normal_customer_address_kh").Value = customerAddressKh
'         ws_print_invoice.Range("print_inv_normal_customer_tel").Value = customerTel & IIF(customerSocialNetwork<>""," , " & customerSocialNetwork,"")
'         ' ws_print_invoice.Range("print_inv_normal_customer_email").Value = customerEmail
'         ' ws_print_invoice.Range("print_inv_normal_customer_tin_number").Value = customerTinNumber
        
'         ws_print_invoice.Range("C27" & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = False 'Show all row in range print invoice

'         If invoiceId > 0 Then
'             k = 0
'             m = 0
'             For j = 1 To countInvoiceDetailItem
'                 If (rangeInvoiceDetail.Cells(j, 2) = invoiceId) Then
'                     productCode = rangeInvoiceDetail.Cells(j, 8)
'                     productName = rangeInvoiceDetail.Cells(j, 9)           'MID(Split(rangeInvoiceDetail.Cells(j, 7),"#")(0), 7, 100)
'                     qty = rangeInvoiceDetail.Cells(j, 10)
'                     qtyFree = rangeInvoiceDetail.Cells(j, 11)
'                     productUoM = Application.VLookup(MID(rangeInvoiceDetail.Cells(j, 12), 1, 3) * 1, rngUom, 5, False)
'                     unitPrice = rangeInvoiceDetail.Cells(j, 18)
'                     totalPrice = rangeInvoiceDetail.Cells(j, 19)

'                     Set criteriaRange1 = ws_inv_service_data.Range("D8:D" & countInvoiceDetailItem + 7) 'column invoice id
'                     Set criteriaRange2 = ws_inv_service_data.Range("AB8:AB" & countInvoiceDetailItem + 7) 'column status
'                     countItem = Application.COUNTIFS(criteriaRange1, invoiceId, criteriaRange2, 1)

'                     If (qty > 0) Then
'                         k = k + 1
'                         m = m + 1
'                         If (countItem >= 26 And countItem <= 104) Then '26, 52, 78, 104
'                             If (k Mod 26 = 0) Then
'                                 k = k + 1
'                             End If
'                         Else
'                             If (k Mod 15 = 0) Then ' k = 13 or k = 26 or k = 39 or k = 52 or k = 65
'                                 k = k + 8
'                             End If
'                         End If
'                         ws_print_invoice.Cells(k + 26,"D") = m
'                         ws_print_invoice.Cells(k + 26,"F") = productCode & " " & productName
'                         ws_print_invoice.Cells(k + 26,"I") = qty
'                         ws_print_invoice.Cells(k + 26,"J") = productUoM
'                         ws_print_invoice.Cells(k + 26,"K") = unitPrice
'                         ws_print_invoice.Cells(k + 26,"M") = totalPrice
'                     End If
'                 End If
'             Next j

'             For i = 1 To countInvoiceServiceItem
'                 If (rangeInvoiceService.Cells(i, 2) = invoiceId) Then
'                     productCode           = rangeInvoiceService.Cells(i, 7)
'                     productName           = rangeInvoiceService.Cells(i, 8) 'MID(Split(rangeInvoiceDetail.Cells(i, 7),"#")(0), 7, 100)
'                     qty                   = rangeInvoiceService.Cells(i, 9)
'                     qtyFree               = rangeInvoiceService.Cells(i, 10)
'                     productUoM            = Application.VLookup(MID(rangeInvoiceService.Cells(i, 11), 1, 3) *1, rngUom,5, False)
'                     unitPrice             = rangeInvoiceService.Cells(i, 17)
'                     totalPrice            = rangeInvoiceService.Cells(i, 18)

'                     Set criteriaRange1 = ws_inv_service_data.Range("D8:D" & countInvoiceServiceItem + 7) 'column invoice id
'                     Set criteriaRange2 = ws_inv_service_data.Range("Y8:Y" & countInvoiceServiceItem + 7) 'column status
'                     countItem = Application.COUNTIFS(criteriaRange1, invoiceId, criteriaRange2, 1)

'                     If (qty > 0) Then
'                         k = k + 1
'                         m = m + 1
'                         If (countItem >= 26 And countItem <= 104) Then '26, 52, 78, 104
'                             If (k Mod 26 = 0) Then
'                                 k = k + 1
'                             End If
'                         Else
'                             If (k Mod 15 = 0) Then ' k = 13 or k = 26 or k = 39 or k = 52 or k = 65
'                                 k = k + 8
'                             End If
'                         End If
'                         ws_print_invoice.Cells(k + 26,"D") = m
'                         ws_print_invoice.Cells(k + 26,"F") = productCode & " " & productName
'                         ws_print_invoice.Cells(k + 26,"I") = qty
'                         ws_print_invoice.Cells(k + 26,"J") = productUoM
'                         ws_print_invoice.Cells(k + 26,"K") = unitPrice
'                         ws_print_invoice.Cells(k + 26,"M") = totalPrice
'                     End If
'                 End If
'             Next i

'             lastRow = k + 27
'             ws_print_invoice.Range("C" & lastRow & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
'             ' OnEnd
'         End If
'     Else
'         rangePrintInvItem.Value = "" ' rangePrintInvItem.ClearContents
'     End If
' End Sub

' Private Sub Worksheet_Activate()
'     ' OnStart
'     Dim activeWorkbook, wbDatabase As Workbook
'     Dim ws_print_invoice, ws_inv_data, ws_file_setting As Worksheet
'     Dim rowInvoiceCode, rowInvoiceId As Range
'     Dim i, k, CountaResultCell, lastInvoiceId As Integer
'     Dim databasePath, databaseFile As String

'     Set activeWorkbook = Workbooks("index.xlsm")
'     Set ws_print_invoice = activeWorkbook.Sheets("PrintInvoiceNormal")
    
'     ws_print_invoice.InvoiceNormalComboBox.Clear
'     ws_print_invoice.InvoiceNormalComboBox.AddItem "Please select invoice code"

'     If (checkDatabaseFile = True) Then
'         Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
'         databasePath = ws_file_setting.Range("database_path").Value
'         databaseFile = databasePath & "database_kscm.xlsm"

'         If (IsWorkbookOpen(databaseFile) = False) Then
'             ' Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
'         End If
        
'         Set wbDatabase = Workbooks("database_kscm.xlsm")
'         Set ws_inv_data = wbDatabase.Sheets("SaleInvoiceData")
'         lastInvoiceId = ws_inv_data.Cells(Rows.Count, "C").end(xlUp).row - 7

'         Set rowInvoiceCode = ws_inv_data.Range("E8:E" & ws_inv_data.Cells(Rows.Count, "C").end(xlUp).row)
'         Set rowInvoiceId = ws_inv_data.Range("C8:C" & ws_inv_data.Cells(Rows.Count, "C").end(xlUp).row)

'         For i = 1 To lastInvoiceId
'             If rowInvoiceCode.Cells(i) <> "" Then
'                 ws_print_invoice.InvoiceNormalComboBox.AddItem rowInvoiceCode(i) & " #" & rowInvoiceId(i)
'             End If
'         Next i
'     Else
'         MsgBox "Database file doesn't exist."
'         Exit Sub
'     End If
'     ' OnEnd
' End Sub

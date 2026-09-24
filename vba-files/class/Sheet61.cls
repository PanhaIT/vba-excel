VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet61"
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

Private Sub PurchaseBillComboBox_Change()
    ' OnStart
    'loop variable
    Dim i, j, k,lastRow As Integer

    'workbook/worksheet variable
    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_gld, ws_gl, ws, ws_inv_data, ws_inv_item_data,ws_file_setting, ws_print_pb As Worksheet

    'vendor variable
    Dim vendorId, vendorName As String
    Dim vendorListRange As Range
    Dim vendorNameEn, vendorNameKh, vendorAddressKh, vendorTel, vendorEmail, vendorTinNumber As String

    'branch variable
    Dim branchId, branchName As String
    Dim branchListRange As Range
    Dim bankName, accountName, accountNo, branchAddressKh, branchTel, branchEmail As String

    'product variable
    Dim productId, qty, qtyFree, productCode, productName, productUoM As String
    Dim rngUom, item_row As Range

    'purchase invoice variable
    Dim pbDate, pbDueDate As Date
    Dim statusNote, priceTypeId, discountTypeId, PbCode, pbNote As String
    Dim purchaseId, netDays, paymentTermId, conversion, qtyOrder, qtyOrderSmall, invValId, pbNo, invId, invDetailId, exchangeRate As Integer
    Dim lastPurchaseId, countPurchaseDetailItem, vatChartAccId, lastPbDetailId As Integer
    Dim rngPurchaseBillList, rowPbCode, rngPbDetail, rangePrintInvItem, rangeInvDetailId  As Range
    Dim criteriaRange1, criteriaRange2 As Range
    Dim countItem As Integer

    'setting variable
    Dim companyId As Integer
    Dim vatPercent, vatSettingId, vatCal, chartAccIncome As Integer
    Dim companyNameKh, companyNameEn,companyTinNumber, databasePath, databaseFile As String
    Dim companyListRange, paymentTermRange As Range

    'amount variable
    Dim totalCost, totalDeposit, totalDiscount, totalAmount, grandTotalAmount, unitCost, totalDisItem, disItem, totalVat As Double
    
    Set activeWorkbook    = Workbooks("index.xlsm")
    Set ws_print_pb       = activeWorkbook.Sheets("PrintPurchaseBill")
    Set rangePrintInvItem = ws_print_pb.Range("print_pb_items")

    If (Me.PurchaseBillComboBox.Value <> "Please select purchase code" And Me.PurchaseBillComboBox.Value <> "") Then 'checkDatabaseFile = TRUE And
        rangePrintInvItem.ClearContents
        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath        = ws_file_setting.Range("database_path").Value
        databaseFile        = databasePath & "database_kscm.xlsm"

        ' If (IsWorkbookOpen(databaseFile) = FALSE) Then
        '     ' Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        '     MsgBox "Please open database file befor select purchase code."
        '     exit sub
        ' End If

        Set wbDatabase       = Workbooks("database_kscm.xlsm")
        Set ws_inv_data      = wbDatabase.Sheets("PurchaseBill")
        Set ws_inv_item_data = wbDatabase.Sheets("PurchaseBillDetail")
        lastPurchaseId       = ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row
        lastPbDetailId       = ws_inv_item_data.Cells(Rows.Count, "C").End(xlUp).Row

        Set rngPurchaseBillList  = ws_inv_data.Range("C8:AP" & lastPurchaseId)
        Set rangeInvDetailId     = ws_inv_item_data.Range("C8:C" & lastPbDetailId)
        Set rngPbDetail          = ws_inv_item_data.Range("C8:Z" & lastPbDetailId)
        
        countPurchaseDetailItem  = Application.CountA(rangeInvDetailId)
        purchaseId               = CInt(Split(Me.PurchaseBillComboBox.Value, "#")(1))
        PbCode                   = Split(Me.PurchaseBillComboBox.Value, "#")(0)
        
        companyId       = 1
        vendorId        = Application.VLookup(purchaseId, rngPurchaseBillList, 14, False)
        pbDate          = Format(Application.VLookup(purchaseId, rngPurchaseBillList, 3, False), "yyyy/mm/dd")
        paymentTermId   = Application.VLookup(purchaseId, rngPurchaseBillList, 13, False)
        branchId        = Application.VLookup(purchaseId, rngPurchaseBillList, 10, False)
        branchName      = ""
        vendorName      = Application.VLookup(purchaseId, rngPurchaseBillList, 15, False)
        pbNote          = Application.VLookup(purchaseId, rngPurchaseBillList, 33, False)

        vatPercent      = Application.VLookup(purchaseId, rngPurchaseBillList, 20, False)
        vatCal          = Application.VLookup(purchaseId, rngPurchaseBillList, 22, False)
        totalVat        = Application.VLookup(purchaseId, rngPurchaseBillList, 26, False)

        exchangeRate    = Application.VLookup(purchaseId, rngPurchaseBillList, 18, False)
        totalAmount     = Application.VLookup(purchaseId, rngPurchaseBillList, 25, False)
        totalDiscount   = Application.VLookup(purchaseId, rngPurchaseBillList, 27, False)
        totalDeposit    = Application.VLookup(purchaseId, rngPurchaseBillList, 28, False)
        grandTotalAmount= Application.VLookup(purchaseId, rngPurchaseBillList, 29, False)
        
        Set companyListRange  = activeWorkbook.Sheets("Setting").Range("company_list")
        Set vendorListRange   = activeWorkbook.Sheets("VendorList").Range("vendor_list")
        Set branchListRange   = activeWorkbook.Sheets("Branch").Range("branch_list")
        Set paymentTermRange  = activeWorkbook.Sheets("Setting").Range("payment_term_list")
        Set rngUom            = activeWorkbook.Sheets("UoM").Range("uom_list")

        netDays   = Application.VLookup(paymentTermId, paymentTermRange, 3, False)
        pbDueDate = 0
        If netDays <> "" Then
            pbDueDate = DateAdd("d", netDays, pbDate)
        End If
        companyNameKh    = Application.VLookup(companyId, companyListRange, 2, False)
        companyNameEn    = Application.VLookup(companyId, companyListRange, 3, False)
        companyTinNumber = Application.VLookup(companyId, companyListRange, 5, False)

        vendorNameEn     = Application.VLookup(vendorId, vendorListRange, 5, False)
        vendorNameKh     = Application.VLookup(vendorId, vendorListRange, 4, False)
        vendorAddressKh  = Application.VLookup(vendorId, vendorListRange, 14, False)
        vendorTel        = Application.VLookup(vendorId, vendorListRange, 8, False)
        vendorEmail      = Application.VLookup(vendorId, vendorListRange, 11, False)
        vendorTinNumber  = Application.VLookup(vendorId, vendorListRange, 12, False)

        bankName         = Split(Application.VLookup(branchId, branchListRange, 6, False), ",")(0)
        accountName      = Split(Application.VLookup(branchId, branchListRange, 6, False), ",")(1)
        accountNo        = Split(Application.VLookup(branchId, branchListRange, 6, False), ",")(2)
        branchAddressKh  = Application.VLookup(branchId, branchListRange, 22, False)
        branchTel        = Application.VLookup(branchId, branchListRange, 18, False)
        branchEmail      = Application.VLookup(branchId, branchListRange, 20, False)
       
        If IsError(companyNameKh) Then
            companyNameKh = ""
        End If

        'pb Information
        ws_print_pb.Range("print_pb_code").Value = PbCode
        ws_print_pb.Range("print_pb_date").Value = pbDate
        ws_print_pb.Range("print_pb_due_date").Value = pbDueDate
        ws_print_pb.Range("print_pb_note").Value = pbNote
        ws_print_pb.Range("print_pb_exchange_rate").Value = exchangeRate
        ws_print_pb.Range("print_pb_company_name_kh").Value = companyNameKh
        ws_print_pb.Range("print_pb_company_name_en").Value = companyNameEn
        ws_print_pb.Range("print_pb_company_tin").Value = companyTinNumber
        ws_print_pb.Range("print_pb_payment_term_id").Value = paymentTermId
        ws_print_pb.Range("print_pb_payment_net_days").Value = netDays
    
        ws_print_pb.Range("print_pb_sub_total").Value = totalAmount
        ws_print_pb.Range("print_pb_sub_total_kh").Value = totalAmount * exchangeRate
        ws_print_pb.Range("print_pb_toal_discount").Value = totalDiscount
        ws_print_pb.Range("print_pb_toal_discount_kh").Value = totalDiscount * exchangeRate
        ws_print_pb.Range("print_pb_toal_deposit").Value = totalDeposit
        ws_print_pb.Range("print_pb_toal_vat").Value = totalVat

        ws_print_pb.Range("print_pb_grand_total").Value = grandTotalAmount
        ws_print_pb.Range("print_pb_grand_total_kh").Value = grandTotalAmount * exchangeRate

        'Branch information bankName,accountName,accountNo,branchAddressKh,branchTel,branchEmail
        ws_print_pb.Range("print_pb_bank_name").Value = bankName
        ws_print_pb.Range("print_pb_account_name").Value = accountName
        ws_print_pb.Range("print_pb_account_no").Value = accountNo
        ws_print_pb.Range("print_pb_branch_address_kh").Value = branchAddressKh
        ws_print_pb.Range("print_pb_branch_telephone").Value = branchTel
        ws_print_pb.Range("print_pb_branch_email").Value = branchEmail

        'vendor Information
        ws_print_pb.Range("print_pb_vendor_name_kh").Value = vendorNameKh
        ws_print_pb.Range("print_pb_vendor_name_en").Value = vendorNameEn
        ws_print_pb.Range("print_pb_vendor_address_kh").Value = vendorAddressKh
        ws_print_pb.Range("print_pb_vendor_tel").Value = vendorTel
        ws_print_pb.Range("print_pb_vendor_email").Value = vendorEmail
        ws_print_pb.Range("print_pb_vendor_tin_number").Value = vendorTinNumber

        ws_print_pb.Range("C27" & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = False 'Show all row in range print invoice
        If purchaseId > 0 And countPurchaseDetailItem > 0 Then
            k = 0
            For j = 1 To countPurchaseDetailItem
                If (rngPbDetail.Cells(j, 2) = purchaseId) Then
                    ' disItem             = rngPbDetail.Cells(j,6)
                    ' totalDisItem        = rngPbDetail.Cells(j,6)
                    productCode     = rngPbDetail.Cells(j, 8)
                    productName     = rngPbDetail.Cells(j, 9)           'MID(Split(rngPbDetail.Cells(j, 7),"#")(0), 7, 100)
                    qty             = rngPbDetail.Cells(j, 10)
                    qtyFree         = rngPbDetail.Cells(j, 11)
                    productUoM      = Application.VLookup(Mid(rngPbDetail.Cells(j, 13), 1, 3) * 1, rngUom, 5, False)
                    unitCost        = rngPbDetail.Cells(j, 20)
                    totalCost       = rngPbDetail.Cells(j, 22)

                    Set criteriaRange1 = ws_inv_item_data.Range("D8:D" & countPurchaseDetailItem + 7) 'column invoice id
                    Set criteriaRange2 = ws_inv_item_data.Range("AB8:AB" & countPurchaseDetailItem + 7) 'column status
                    countItem          = Application.COUNTIFS(criteriaRange1,purchaseId,criteriaRange2,1)

                    If (qty > 0) Then
                        k = k + 1
                        'set condition to manage page 
                        ' If (countItem >= 26 AND countItem <= 104) Then '26, 52, 78, 104
                        '     If (k Mod 26 = 0) Then  '=> limit 26 rows per page
                        '         ' Debug.Print "***************k=" & k
                        '         k = k + 1
                        '     End If
                        ' Else
                        '     If (k Mod 13 = 0) Then ' k = 13 => limit 12 rows per page
                        '         ' Debug.Print "***************k=" & k
                        '         k = k + 14
                        '     End If
                        ' End If
                        ' debug.Print "k=" & k
                        rangePrintInvItem.Cells(k, 1) = productCode
                        rangePrintInvItem.Cells(k, 2) = productName
                        rangePrintInvItem.Cells(k, 5) = productUoM
                        rangePrintInvItem.Cells(k, 6) = qty
                        rangePrintInvItem.Cells(k, 7) = unitCost
                        rangePrintInvItem.Cells(k, 8) = totalCost
                        rangePrintInvItem.Cells(k, 9) = totalCost * exchangeRate
                    End If
                End If
            Next j
            lastRow = k + 27
            ws_print_pb.Range("C" & lastRow & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
        End If
    Else
        ' MsgBox "Database file doesn't exist."
        ' Exit Sub
        rangePrintInvItem.ClearContents
        ws_print_pb.Range("C27:C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
    End If

    ' OnEnd
End Sub

Private Sub Worksheet_Activate()
    ' OnStart

    ' ' pb Information
    ' ws_print_pb.Range("print_pb_code").Value = ""
    ' ws_print_pb.Range("print_pb_date").Value = ""
    ' ws_print_pb.Range("print_pb_due_date").Value = ""
    ' ws_print_pb.Range("print_pb_note").Value = ""
    ' ws_print_pb.Range("print_pb_exchange_rate").Value = ""
    ' ws_print_pb.Range("print_pb_company_name_kh").Value = ""
    ' ws_print_pb.Range("print_pb_company_name_en").Value = ""
    ' ws_print_pb.Range("print_pb_company_tin").Value = ""
    ' ws_print_pb.Range("print_pb_payment_term_id").Value = ""
    ' ws_print_pb.Range("print_pb_payment_net_days").Value = ""

    ' ws_print_pb.Range("print_pb_sub_total").Value = 0
    ' ws_print_pb.Range("print_pb_sub_total_kh").Value = 0
    ' ws_print_pb.Range("print_pb_toal_discount").Value = 0
    ' ws_print_pb.Range("print_pb_toal_discount_kh").Value = 0
    ' ws_print_pb.Range("print_pb_toal_deposit").Value = 0
    ' ws_print_pb.Range("print_pb_toal_vat").Value = 0
    ' ws_print_pb.Range("print_pb_grand_total").Value = 0
    ' ws_print_pb.Range("print_pb_grand_total_kh").Value = 0

    ' ' Branch information bankName,accountName,accountNo,branchAddressKh,branchTel,branchEmail
    ' ws_print_pb.Range("print_pb_bank_name").Value = ""
    ' ws_print_pb.Range("print_pb_account_name").Value = ""
    ' ws_print_pb.Range("print_pb_account_no").Value = ""
    ' ws_print_pb.Range("print_pb_branch_address_kh").Value = ""
    ' ws_print_pb.Range("print_pb_branch_telephone").Value = ""
    ' ws_print_pb.Range("print_pb_branch_email").Value = ""

    ' ' ' vendor Information
    ' ws_print_pb.Range("print_pb_vendor_name_kh").Value = ""
    ' ws_print_pb.Range("print_pb_vendor_name_en").Value = ""
    ' ws_print_pb.Range("print_pb_vendor_address_kh").Value = ""
    ' ws_print_pb.Range("print_pb_vendor_tel").Value = ""
    ' ws_print_pb.Range("print_pb_vendor_email").Value = ""
    ' ws_print_pb.Range("print_pb_vendor_tin_number").Value = ""

    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_print_pb, ws_inv_data, ws_file_setting As Worksheet
    Dim rowPbCode, rowpurchaseId As Range
    Dim i, CountaResultCell, lastPurchaseId As Integer
    Dim databasePath, databaseFile As String

    Set activeWorkbook = Workbooks("index.xlsm")
    Set ws_print_pb = activeWorkbook.Sheets("PrintPurchaseBill")
    
    ws_print_pb.PurchaseBillComboBox.Clear
    ws_print_pb.PurchaseBillComboBox.AddItem "Please select purchase code"

    If (checkDatabaseFile = True) Then
       
        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath = ws_file_setting.Range("database_path").Value
        databaseFile = databasePath & "database_kscm.xlsm"

        If (IsWorkbookOpen(databaseFile) = False) Then
            ' Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        End If
        Set wbDatabase  = Workbooks("database_kscm.xlsm")
        Set ws_inv_data = wbDatabase.Sheets("PurchaseBill")
        lastPurchaseId  = ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row - 7
    
        Set rowPbCode = ws_inv_data.Range("D8:D" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row)
        Set rowpurchaseId = ws_inv_data.Range("C8:C" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row)
        For i = 1 To lastPurchaseId
            If rowPbCode.Cells(i) <> "" Then
                ws_print_pb.PurchaseBillComboBox.AddItem rowPbCode(i) & " #" & rowpurchaseId(i)
            End If
        Next i
    Else
        ' MsgBox "Database file doesn't exist."
        ' Exit Sub
    End If
    
    ' OnEnd
End Sub

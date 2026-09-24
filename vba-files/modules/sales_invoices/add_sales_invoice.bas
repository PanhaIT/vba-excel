Attribute VB_Name = "add_sales_invoice"

Private Sub SaveSaleInvoice()
    OnStart
    Dim ws_gld, ws_gl,ws_invoice_data, ws_add_inv As Worksheet
    Dim unitPrice, newUnitPrice,unitCostByUom,unitPriceBigUom, totalPrice, totalDeposit,totalDiscount,totalAmount,grandTotalAmount,unitCostBigUom,totalDisItem,disItem,totalVat As Double
    Dim statusNote,customerId,invoiceSheet,branchId,locationId,priceTypeId,discountTypeId,saleId,saleName As string
    Dim customerName,salesInvoiceNote,invoiceWeek,invoiceMonth,invoiceYear,expenseTypeId,expenseType,branchName,msgSave As String
    Dim conversion, qtyOrder,qtyOrderSmall,invValId,statusInvoice, invoiceNo,taxInvoiceCode, indexColor, invDetailId, inventoryId, invTotalDetailId, invTotalId,itemType,invoiceType As Integer

    Dim qtyInvAdj, qtySale, qtySaleFree, qtyPb, qtyPr, qtySr, qtySrFree, qtyToIn, qtyToOut,stockIn,stockOut,totalQty,coaSalesVat,gld_sid_id,gld_cogs_id,gld_inventory_id As Integer
    Dim paymentTermId, qtyStock, productId, qty, qtyFree,convertion,smallValUom, checkExistItemInvTotal, checkExistItemInvTotalDetail,vatPercent,vatSettingId,vatCal,coaSaleIncome,gld_sales_dis_id As Integer
    Dim productCode, productName, productUoM,invoiceCode,lotsNumber As String

    Dim rngProductList, item_row, rd_g, si_data_r, invTotal,ws_gld_data,ws_gl_data, invValuation As Range
    Dim i, j, a, b, LastRowInvTotal, lastRowInvTotalDetail  As Long
    Dim invTotalGroupDetail,invGroupTotal, ws_inv, ws_invTotal, ws_inv_total_detail,ws_invVal,ws_group_total_detail,ws_group_total As Worksheet
    Dim gld_discount_id,lastInvTotalGroupDetailId, LastRowInvGroupTotal, invTotalGroupDetailId,invGroupTotalId, checkExistItemInvGroupTotal,checkExistItemInvGroupTotalDetail,stockByDateInvGroupTotalDetail,stockAvailableInvGroupTotal As Integer
    Dim coaDefaultRange As Range
    Dim coaSaleDiscount,last_row_deposit, last_row_vat, last_row_total_dis As Integer
    Dim totalCost, totalPriceBeforDis As Double
    Dim  invId As Long
    
    Dim invoiceDate, createdDate,expiredDate As Date
    Dim confirmBoxAlert As VbMsgBoxResult
    Dim dataInputPath As String
    
    msgSave           = "Are you want to save this invoice"
    confirmBoxAlert   = MsgBoxW(msgSave, vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Add Sales Invoice")
    'taxInvoiceCode   = ActiveSheet.Range("invoice_tax_code").Value
    invoiceNo         = ActiveSheet.Range("invoice_no").Value
    invoiceCode       = ActiveSheet.Range("invoice_code").Value
    customerId        = ActiveSheet.Range("invoice_customer_id").Value
    invoiceDate       = ActiveSheet.Range("invoice_date").Value 'Format(ActiveSheet.Range("invoice_date").Value, "dd/mm/yyyy")
    branchId          = ActiveSheet.Range("invoice_branch_id").Value
    
    If (ActiveSheet.Range("invoice_branch").Value <> "") Then 
        branchName        = MID(ActiveSheet.Range("invoice_branch").Value,5,100)
    Else
        MsgBox "Branch is requied, try again."
        Exit Sub
    End If

    locationGroupId   = ActiveSheet.Range("invoice_location_group_id").Value
    locationId        = ActiveSheet.Range("invoice_location_id").Value
    priceTypeId       = ActiveSheet.Range("invoice_price_type_id").Value
    discountTypeId    = ActiveSheet.Range("invoice_discount_id").Value
    saleId            = ActiveSheet.Range("invoice_seller_id").Value

    If (ActiveSheet.Range("invoice_seller").Value <> "") Then 
        saleName      = MID(ActiveSheet.Range("invoice_seller").Value,5,100)
    Else
        saleName      = "General"
    End If
    
    If (ActiveSheet.Range("invoice_customer").Value <> "") Then 
        customerName      = MID(ActiveSheet.Range("invoice_customer").Value,6,100)
    Else
        MsgBox "Customer is requied, try again."
        Exit Sub
    End If

    salesInvoiceNote  = ActiveSheet.Range("invoice_note").Value
    invoiceType       = ActiveSheet.Range("invoice_type_id").Value
    paymentTermId     = ActiveSheet.Range("invoice_payment_term_id").Value
    currencyId        = ActiveSheet.Range("invoice_currency_id").Value

    If (ActiveSheet.Range("invoice_date").Value > 0) Then 
        invoiceMonth  = month(ActiveSheet.Range("invoice_date").Value)
        invoiceYear   = year(ActiveSheet.Range("invoice_date").Value)
    Else
        MsgBox "Invoice date is requied, try again."
        Exit Sub
    End If

    invoiceWeek       = ActiveSheet.Range("invoice_week").Value
    createdDate       = Format(Now(), "yyyy/mm/dd hh:mm:ss")

    vatPercent        = ActiveSheet.Range("invoice_vat_percent").Value
    If (ActiveSheet.Range("invoice_vat").Value <> "") Then 
        vatSettingId  = MID(ActiveSheet.Range("invoice_vat").Value,1,2)*1
    Else
        vatSettingId  = 1
    End If
    
    vatCal            = ActiveSheet.Range("invoice_vat_cal_id").Value
    totalVat          = Format(ActiveSheet.Range("invoice_total_vat").Value,"0.00")

    totalDeposit      = Format(ActiveSheet.Range("invoice_deposit").Value,"0.00")
    totalDiscount     = Format(ActiveSheet.Range("invoice_discount").Value,"0.00")
    totalAmount       = Format(ActiveSheet.Range("invoice_sub_total_amount").Value,"0.00")
    statusInvoice     = Format(ActiveSheet.Range("invoice_status").Value,"0.00")
    grandTotalAmount  = Format(ActiveSheet.Range("invoice_total_amount").Value,"0.00")

    Set coaDefaultRange = Sheets("ChartAccountDefault").Range("chart_account_default_list2")
    coaSaleDiscount     = Application.VLookup("Sales Discount", coaDefaultRange, 3, False)
    coaSaleIncome       = Application.VLookup("Sales Income", coaDefaultRange, 3, False)
    coaReceivable       = Application.VLookup("Accounts Receivable", coaDefaultRange, 3, False)
    coaSalesVat         = Application.VLookup("Sales VAT", coaDefaultRange, 3, False)

    If (ActiveSheet.Range("invoice_exchange_rate").Value >0) Then 
        exchangeRate  = ActiveSheet.Range("invoice_exchange_rate").Value
    Else
        exchangeRate  = 4000
    End If
    totalDeposit      = 0
    totalBalance      = totalAmount + totalVat - (totalDiscount + totalDeposit)

    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_product, ws_file_setting As Worksheet
    Dim MyFSO As New FileSystemObject
    Dim databasePath, databaseFile As String

    Set activeWorkbook = Workbooks("index.xlsm")
    Set ws_add_inv     = activeWorkbook.Sheets("AddInvoice")
    Set ws_product     = activeWorkbook.Sheets("ProductList")
    Set rngProductList = ws_product.Range("C8:AA" & ws_product.Cells(Rows.Count,"C").End(xlUp).Row)

    If confirmBoxAlert = vbYes And checkDatabaseFile = TRUE  Then

        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath        = ws_file_setting.Range("database_path").Value
        databaseFile        = databasePath & "database_kscm.xlsm"
        dataInputPath       = ws_file_setting.Range("data_input_path").value

        Dim check_access As Integer
        check_access = 0 
        If (IsWorkbookOpen(CStr(databaseFile)) = FALSE) Then
            Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
            check_access = 1
        End If

        If (check_access = 1) Then
            Run resetAllFilter()

            'Sales Invoice Status=> -1 = Edit, 1  = issue, 2  = fulfilled ,3  = partial
            ' statusInvoice = -1
            ' indexColor    = 1
            ' statusNote    = "Edit"
            If totalBalance = 0 Then
                statusInvoice = 2
                indexColor = 43
                statusNote = "fulfilled"
            Elseif totalBalance > 0 Then
                If totalDeposit>0 AND  totalDeposit < (totalAmount + totalVat - totalDiscount) Then 
                    statusInvoice = 3
                    indexColor = 44
                    statusNote = "Partial"
                ElseIf totalDeposit = 0 Then  
                    statusInvoice = 1
                    indexColor = 3
                    statusNote = "Issue"
                End If
            End If
            statusInvoice = statusInvoice * 1

            If statusInvoice > 0 Then
                If invoiceDate = "" Or priceTypeId = "" Or customerId = "" Or totalAmount = 0 Then
                    MsgBox "Please Select require field, try again."
                    Exit Sub
                Else
                    Dim access As Integer
                    access = 0
                    If totalAmount > 0 Then
                        '***Start save date To sales invoice data***
                        Dim gl_id,gld_id As Integer
                        Set ws_invoice_data        = wbDatabase.Sheets("SaleInvoiceData")
                        Set ws_inv_detail          = wbDatabase.Sheets("SalesByItemData")

                        Set ws_gl                  = wbDatabase.sheets("GeneralLedger")
                        Set ws_gld                 = wbDatabase.Sheets("GeneralLedgerDetail")

                        Set ws_inv                 = wbDatabase.Sheets("inventories")
                        Set ws_invVal              = wbDatabase.Sheets("InventoryValuation")

                        Set ws_invTotal            = wbDatabase.Sheets(locationId & "_inventory_totals")
                        Set ws_inv_total_detail    = wbDatabase.Sheets(locationId & "_inventory_total_details")

                        Set ws_group_total         = wbDatabase.Sheets(locationGroupId & "_group_totals")
                        Set ws_group_total_detail  = wbDatabase.Sheets(locationGroupId & "_group_total_details")

                        Set si_data_r        = ws_invoice_data.Range("C8:AI8")
                        invId                = ws_invoice_data.Cells(Rows.Count, "C").End(xlUp).Row - 6

                        Set rd_g             = ws_inv_detail.Range("C8:AA8")
                        invDetailId          = ws_inv_detail.Cells(Rows.Count, "C").End(xlUp).Row - 6

                        Set item_row         = activeWorkbook.Sheets("AddInvoice").Range("I8:AM87")

                        Set ws_gl_data       = ws_gl.Range("C8:Al8")
                        gl_id                = ws_gl.Cells(Rows.Count, "C").End(xlUp).Row - 6 'gl = general ledger

                        Set ws_gld_data      = ws_gld.Range("C8:AC8")
                        gld_id               = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6 'gld = general ledger detail

                        Dim m As Integer

                        '*************Start run tax invoice code******************
                        m = 1
                        For k = 1 To ws_invoice_data.Cells(Rows.Count, "C").End(xlUp).Row - 7
                            If (si_data_r.Cells(k,28) > 0 And si_data_r.Cells(k,4) < invoiceDate) Then
                                If (si_data_r.Cells(k,2) <> si_data_r.Cells(k+1,2))  Then
                                    m = m + 1
                                End If
                            End If
                        Next k
  
                        Dim rngInvoiceType As Range
                        Dim invPrefix As String
                        Set rngInvoiceType = activeWorkbook.Sheets("Setting").Range("invoice_type_list")
                        invPrefix          = Application.Vlookup(CInt(invoiceType), rngInvoiceType,3, False)
                        taxInvoiceCode     = invPrefix & Mid(invoiceYear,3,4) & "-" & PadStr(m + 54, 4, "0", xlHAlignRight) 'generate tax invoice code

                        activeWorkbook.Sheets("AddInvoice").Range("invoice_tax_no").Value   = m
                        activeWorkbook.Sheets("AddInvoice").Range("invoice_tax_code").Value = taxInvoiceCode
                        '*************End run tax invoice code******************

                        'insert sale invoice
                        If (totalVat<=0) Then 
                            vatCal          = ""
                            vatSettingId    = ""
                            vatPercent      = 0
                            coaSalesVat   = ""
                        End If

                        si_data_r.Cells(invId,1)  = invId 'id
                        si_data_r.Cells(invId,2)  = taxInvoiceCode 'tax invoice no.
                        si_data_r.Cells(invId,3)  = invoiceCode 'invoice no.
                        si_data_r.Cells(invId,4)  = invoiceDate 'invoice date
                        si_data_r.Cells(invId,5)  = createdDate 'created date
                        si_data_r.Cells(invId,6)  = branchId 'branch id
                        si_data_r.Cells(invId,7)  = locationGroupId 'location group id
                        si_data_r.Cells(invId,8)  = locationId 'location id
                        si_data_r.Cells(invId,9)  = priceTypeId 'price type id
                        si_data_r.Cells(invId,10) = discountTypeId 'discount type id
                        si_data_r.Cells(invId,11) = saleId 'sale id
                        si_data_r.Cells(invId,12) = customerId 'customer id
                        si_data_r.Cells(invId,13) = paymentTermId 'payment term id
                        si_data_r.Cells(invId,14) = currencyId 'currency id 
                        si_data_r.Cells(invId,15) = coaReceivable 'chart account ar id
                        si_data_r.Cells(invId,16) = saleName 'sale name
                        si_data_r.Cells(invId,17) = customerName 'customer name
                        si_data_r.Cells(invId,18) = exchangeRate 'exhange rate
                        si_data_r.Cells(invId,19) = coaSalesVat 'Vat Chart Account Id
                        si_data_r.Cells(invId,20) = vatPercent 'Vat Percent
                        si_data_r.Cells(invId,21) = vatSettingId 'Vat Setting Id
                        si_data_r.Cells(invId,22) = vatCal '  Vat Cal
                        si_data_r.Cells(invId,23) = totalAmount 'total amount
                        si_data_r.Cells(invId,24) = totalVat 'total vat
                        si_data_r.Cells(invId,25) = totalDiscount 'total discount
                        si_data_r.Cells(invId,26) = totalDeposit 'total deposit
                        si_data_r.Cells(invId,27) = totalBalance 'balance
                        si_data_r.Cells(invId,28) = statusInvoice 'sale invoice status
                        si_data_r.Cells(invId,29) = salesInvoiceNote 'sale invoice note
                        si_data_r.Cells(invId,30) = invoiceType 'invoice type
                        si_data_r.Cells(invId,31) = invoiceWeek 'week
                        si_data_r.Cells(invId,32) = invoiceMonth 'month
                        si_data_r.Cells(invId,33) = invoiceYear 'year
                        Call setStatusColor(si_data_r.Cells(invId,28))
                        '***End save date To sales invoice data***

                        'insert general ledger
                        ws_gl_data.Cells(gl_id,1)  = gl_id 'id
                        ws_gl_data.Cells(gl_id,3)  = invId 'sales_invoice_id
                        ws_gl_data.Cells(gl_id,23) = invoiceDate 'date
                        ws_gl_data.Cells(gl_id,24) = invoiceCode 'reference
                        ws_gl_data.Cells(gl_id,25) = 0 'total_deposit
                        ws_gl_data.Cells(gl_id,27) = createdDate 'created
                        ws_gl_data.Cells(gl_id,28) = 1 'created_by
                        ws_gl_data.Cells(gl_id,31) = 1 'is_approve
                        ws_gl_data.Cells(gl_id,32) = 0 'is_depreciated
                        ws_gl_data.Cells(gl_id,33) = 0 'is_retained_earnings
                        ws_gl_data.Cells(gl_id,34) = 0 'deposit_type
                        ws_gl_data.Cells(gl_id,35) = 1 'is_active
                        ws_gl_data.Cells(gl_id,36) = invoiceWeek 'week
                        ws_gl_data.Cells(gl_id,37) = invoiceMonth 'month
                        ws_gl_data.Cells(gl_id,38) = invoiceYear 'year

                        'insert general ledger detail Accounts Receivable
                        ws_gld_data.Cells(gld_id,1)  = gld_id 'id
                        ws_gld_data.Cells(gld_id,2)  = gl_id 'general_ledger_id
                        ws_gld_data.Cells(gld_id,3)  = coaReceivable 'chart_account_id
                        ws_gld_data.Cells(gld_id,4)  = 1 'company_id
                        ws_gld_data.Cells(gld_id,5)  = branchId 'branch id
                        ws_gld_data.Cells(gld_id,6)  = locationGroupId 'location group id
                        ws_gld_data.Cells(gld_id,7)  = locationId 'location id
                        ws_gld_data.Cells(gld_id,16) = "Invoice" 'type
                        ws_gld_data.Cells(gld_id,17) = totalBalance 'debit
                        ws_gld_data.Cells(gld_id,18) = 0 'credit
                        ws_gld_data.Cells(gld_id,19) = "ICS: INV # " & invoiceCode 'memo
                        ws_gld_data.Cells(gld_id,20) = customerId 'customer_id
                        ws_gld_data.Cells(gld_id,24) = 1 'class_id
                        ws_gld_data.Cells(gld_id,25) = 1 'is_active
                        ws_gld_data.Cells(gld_id,26) = invoiceWeek 'week
                        ws_gld_data.Cells(gld_id,27) = invoiceMonth 'month
                        ws_gld_data.Cells(gld_id,28) = invoiceYear 'year
                        
                        '*******Start save invoice discount*******
                        If totalDiscount > 0 Then
                            'insert general ledger detail
                            last_row_total_dis = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                            ws_gld_data.Cells(last_row_total_dis,1)  = last_row_total_dis 'id
                            ws_gld_data.Cells(last_row_total_dis,2)  = gl_id 'general_ledger_id
                            ws_gld_data.Cells(last_row_total_dis,3)  = 12 'chart_account_id application.VLOOKUP("Sales Discount",chart_account_default_list2,3,FALSE)
                            ws_gld_data.Cells(last_row_total_dis,4)  = 1 'company_id
                            ws_gld_data.Cells(last_row_total_dis,5)  = branchId 'branch id
                            ws_gld_data.Cells(last_row_total_dis,6)  = locationGroupId 'location group id
                            ws_gld_data.Cells(last_row_total_dis,7)  = locationId 'location id
                            ws_gld_data.Cells(last_row_total_dis,16) = "Invoice" 'type
                            ws_gld_data.Cells(last_row_total_dis,17) = totalDiscount 'debit
                            ws_gld_data.Cells(last_row_total_dis,18) = 0 'credit
                            ws_gld_data.Cells(last_row_total_dis,19) = "ICS: INV # " & invoiceCode & " Total Discount" 'memo
                            ws_gld_data.Cells(last_row_total_dis,20) = customerId 'customer_id
                            ws_gld_data.Cells(last_row_total_dis,24) = 1 'class_id
                            ws_gld_data.Cells(last_row_total_dis,25) = 1 'is_active
                            ws_gld_data.Cells(last_row_total_dis,26) = invoiceWeek 'week
                            ws_gld_data.Cells(last_row_total_dis,27) = invoiceMonth 'month
                            ws_gld_data.Cells(last_row_total_dis,28) = invoiceYear 'year
                            last_row_total_dis = last_row_total_dis + 1
                        End If
                        '*******End save invoice discount*******

                        '*******Start save invoice vat*******
                        If totalVat > 0 Then
                            'insert general ledger detail
                            last_row_vat = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                            ws_gld_data.Cells(last_row_vat,1)  = last_row_vat 'id
                            ws_gld_data.Cells(last_row_vat,2)  = gl_id 'general_ledger_id
                            ws_gld_data.Cells(last_row_vat,3)  = coaSalesVat 'chart_account_id
                            ws_gld_data.Cells(last_row_vat,4)  = 1 'company_id
                            ws_gld_data.Cells(last_row_vat,5)  = branchId 'branch id
                            ws_gld_data.Cells(last_row_vat,6)  = locationGroupId 'location group id
                            ws_gld_data.Cells(last_row_vat,7)  = locationId 'location id
                            ws_gld_data.Cells(last_row_vat,16) = "Invoice" 'type
                            ws_gld_data.Cells(last_row_vat,17) = 0 'debit
                            ws_gld_data.Cells(last_row_vat,18) = totalVat 'credit
                            ws_gld_data.Cells(last_row_vat,19) = "ICS: INV # " & invoiceCode & " Total VAT" 'memo
                            ws_gld_data.Cells(last_row_vat,20) = customerId 'customer_id
                            ws_gld_data.Cells(last_row_vat,24) = 1 'class_id
                            ws_gld_data.Cells(last_row_vat,25) = 1 'is_active
                            ws_gld_data.Cells(last_row_vat,26) = invoiceWeek 'week
                            ws_gld_data.Cells(last_row_vat,27) = invoiceMonth 'month
                            ws_gld_data.Cells(last_row_vat,28) = invoiceYear 'year
                            last_row_vat = last_row_vat + 1
                        End If
                        '*******End save invoice vat*******

                        '*******Start save invoice deposit*******
                        If totalDeposit > 0 Then
                            'insert general ledger detail
                            last_row_deposit = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                            ws_gld_data.Cells(last_row_deposit,1)  = last_row_deposit 'id
                            ws_gld_data.Cells(last_row_deposit,2)  = gl_id 'general_ledger_id
                            ws_gld_data.Cells(last_row_deposit,3)  = 2 'chart_account_id AR
                            ws_gld_data.Cells(last_row_deposit,4)  = 1 'company_id
                            ws_gld_data.Cells(last_row_deposit,5)  = branchId 'branch id
                            ws_gld_data.Cells(last_row_deposit,6)  = locationGroupId 'location group id
                            ws_gld_data.Cells(last_row_deposit,7)  = locationId 'location id
                            ws_gld_data.Cells(last_row_deposit,16) = "Invoice" 'type
                            ws_gld_data.Cells(last_row_deposit,17) = totalDeposit 'debit
                            ws_gld_data.Cells(last_row_deposit,18) = 0 'credit
                            ws_gld_data.Cells(last_row_deposit,19) = "ICS: INV # " & invoiceCode & " Total Deposit" 'memo
                            ws_gld_data.Cells(last_row_deposit,20) = customerId 'customer_id
                            ws_gld_data.Cells(last_row_deposit,24) = 1 'class_id
                            ws_gld_data.Cells(last_row_deposit,25) = 1 'is_active
                            ws_gld_data.Cells(last_row_deposit,26) = invoiceWeek 'week
                            ws_gld_data.Cells(last_row_deposit,27) = invoiceMonth 'month
                            ws_gld_data.Cells(last_row_deposit,28) = invoiceYear 'year
                            last_row_deposit = last_row_deposit + 1
                        End If
                        '*******End save invoice deposit*******

                        For a = 1 To item_row.Rows.Count
                            If item_row.Cells(a,1) <> "" Then 'old condition => WorksheetFunction.Count(rng.Rows(a)) <> 0
                                'insert sales invoice detail 
                                conversion          = 1
                                productId           = item_row.Cells(a,1)
                                productCode         = item_row.Cells(a,3)
                                productName         = Application.VLookup(productId * 1, rngProductList, 6, False) 'MID(Split(item_row.Cells(a,5),"#")(0),7,100)
                                qty                 = item_row.Cells(a, 6) * 1
                                qtyFree             = item_row.Cells(a, 7) * 1
                                productUoM          = item_row.Cells(a,8)
                                lotsNumber          = item_row.Cells(a,9)
                                expiredDate         = item_row.Cells(a,10)
                                newUnitPrice        = item_row.Cells(a,12)
                                unitPriceByUom      = item_row.Cells(a,13)
                                disItem             = Format(item_row.Cells(a,15),"0.000")
                                totalPrice          = Format(item_row.Cells(a,17),"0.000")
                                remark              = item_row.Cells(a,18)
                                If (item_row.Cells(a,19)>0) Then
                                    conversion      = item_row.Cells(a,19)
                                End If
                                smallValUom         = item_row.Cells(a,20)
                                totalDisItem        = Format(item_row.Cells(a,21),"0.000")
                                unitCostByUom       = Format(item_row.Cells(a,22),"0.0000")
                                ' unitCostBigUom      = Format(item_row.Cells(a,23),"0.00")
                                unitPriceBigUom     = Format(item_row.Cells(a,24),"0.0000")
                                itemType            = item_row.Cells(a,25) '1 : product, 2:service

                                If (newUnitPrice > 0 ) Then 
                                    unitPrice = Format(newUnitPrice, "0.0000")
                                Else
                                    unitPrice = Format(unitPriceByUom, "0.0000")
                                End If
                                totalCost          = (qty + qtyFree) * unitCostByUom
                                totalPriceBeforDis = Format(qty * unitPrice, "0.000")

                                checkExistItemInvTotal             = item_row.Cells(a,26)
                                checkExistItemInvTotalDetail       = item_row.Cells(a,27)
                                checkExistItemInvGroupTotal        = item_row.Cells(a,28)
                                checkExistItemInvGroupTotalDetail  = item_row.Cells(a,29)
                                stockByDateInvGroupTotalDetail     = item_row.Cells(a,30)
                                stockAvailableInvGroupTotal        = item_row.Cells(a,31)
                                
                                qtyOrder        = (qty + qtyFree) / conversion '(qty + qtyFree) / smallValUom / convertion
                                qtyOrderSmall   = (qty + qtyFree) * (smallValUom / conversion)

                                ' debug.Print "totalPrice=" & totalPrice & " ,disItem=" &  disItem & " ,totalDisItem=" &  totalDisItem & " ,totalPriceBeforDis=" & totalPriceBeforDis

                                rd_g.Cells(invDetailId, 1)   = invDetailId 'id
                                rd_g.Cells(invDetailId, 2)   = invId 'invoice id
                                rd_g.Cells(invDetailId, 3)   = "" 'invoice service id
                                rd_g.Cells(invDetailId, 4)   = taxInvoiceCode 'invoice number
                                rd_g.Cells(invDetailId, 5)   = invoiceCode 'invoice number
                                rd_g.Cells(invDetailId, 6)   = invoiceDate 'invoice date
                                rd_g.Cells(invDetailId, 7)   = customerId 'customer id
                                rd_g.Cells(invDetailId, 8)   = productCode 'barcode product
                                rd_g.Cells(invDetailId, 9)   = productName 'item description/service
                                rd_g.Cells(invDetailId, 10)   = qty 'qty
                                rd_g.Cells(invDetailId, 11)  = qtyFree 'qty free
                                rd_g.Cells(invDetailId, 12)  = productUoM 'uom
                                rd_g.Cells(invDetailId, 13)  = conversion 'uom smallValUom
                                rd_g.Cells(invDetailId, 14)  = "" 'Dis ID
                                rd_g.Cells(invDetailId, 15)  = "" 'Dis Percent
                                rd_g.Cells(invDetailId, 16)  = disItem 'Dis Amount
                                rd_g.Cells(invDetailId, 17)  = unitCostByUom 'unit cost
                                rd_g.Cells(invDetailId, 18)  = unitPrice 'unit price
                                rd_g.Cells(invDetailId, 19)  = totalPrice 'total price
                                rd_g.Cells(invDetailId, 20)  = lotsNumber 'lot number
                                If (expiredDate > 0) Then 
                                rd_g.Cells(invDetailId, 21)  = expiredDate 'expired date
                                Else
                                rd_g.Cells(invDetailId, 21)  = "" 'expired date
                                End If
                                rd_g.Cells(invDetailId, 22)  = remark 'remark
                                rd_g.Cells(invDetailId, 23)  = invoiceWeek 'week
                                rd_g.Cells(invDetailId, 24)  = invoiceMonth 'month
                                rd_g.Cells(invDetailId, 25)  = invoiceYear 'year
                                rd_g.Cells(invDetailId, 26)  = 1 'status

                                'insert inventory valuation
                                Set invValuation     = ws_invVal.Range("C8:AI8")
                                invValId             = ws_invVal.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                
                                invValuation.Cells(invValId, 1)    = invValId 'id
                                invValuation.Cells(invValId, 2)    = 1 'company_id
                                invValuation.Cells(invValId, 3)    = branchId 'branch_id
                                invValuation.Cells(invValId, 5)    = invId 'sales invoice id
                                invValuation.Cells(invValId, 13)   = "Invoice" 'type=> Bill,Inventory Adjust,Invoice
                                invValuation.Cells(invValId, 14)   = invoiceCode 'reference
                                invValuation.Cells(invValId, 15)   = customerId 'customer_id
                                invValuation.Cells(invValId, 16)   = "" 'vendor_id
                                invValuation.Cells(invValId, 17)   = invoiceDate 'date
                                invValuation.Cells(invValId, 18)   = productId 'product_id
                                invValuation.Cells(invValId, 19)   = qtyOrderSmall * -1'small_qty
                                invValuation.Cells(invValId, 20)   = qtyOrder * -1'qty
                                invValuation.Cells(invValId, 21)   = unitCostByUom 'cost
                                invValuation.Cells(invValId, 22)   = 0 'price
                                invValuation.Cells(invValId, 23)   = 0 'on_hand
                                invValuation.Cells(invValId, 24)   = 0 'on_hand_small
                                invValuation.Cells(invValId, 25)   = unitCostByUom 'avg_cost
                                invValuation.Cells(invValId, 26)   = 0 'asset_value
                                invValuation.Cells(invValId, 27)   = createdDate 'created
                                invValuation.Cells(invValId, 28)   = "" 'date_edited
                                invValuation.Cells(invValId, 29)   = 0 'is_refer_gm_id
                                invValuation.Cells(invValId, 30)   = 0 'avg_refer
                                invValuation.Cells(invValId, 31)   = 1 'is_var_cost
                                invValuation.Cells(invValId, 32)   = 0 'is_adjust_value
                                invValuation.Cells(invValId, 33)   = 1 'is_active

                                'insert general ledger detail COGS
                                If (totalCost > 0) Then
                                    'General Ledger Detail (Inventory)
                                    gld_inventory_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                    ws_gld_data.Cells(gld_inventory_id,1)  = gld_inventory_id 'id
                                    ws_gld_data.Cells(gld_inventory_id,2)  = gl_id 'general_ledger_id
                                    ws_gld_data.Cells(gld_inventory_id,3)  = 3 'chart_account_id application.VLOOKUP("Sales Income",chart_account_default_list2,3,FALSE)
                                    ws_gld_data.Cells(gld_inventory_id,4)  = 1 'company_id
                                    ws_gld_data.Cells(gld_inventory_id,5)  = branchId 'branch id
                                    ws_gld_data.Cells(gld_inventory_id,6)  = locationGroupId 'location group id
                                    ws_gld_data.Cells(gld_inventory_id,7)  = locationId 'location id
                                    ws_gld_data.Cells(gld_inventory_id,8)  = productId 'product_id
                                    ws_gld_data.Cells(gld_inventory_id,11) = invValId 'inventory_valuation_id
                                    ws_gld_data.Cells(gld_inventory_id,12) = 0 'inventory_valuation_is_debit
                                    ws_gld_data.Cells(gld_inventory_id,14) = invDetailId 'sales_invoice_detail_id
                                    ws_gld_data.Cells(gld_inventory_id,16) = "Invoice" 'type=> Bill,Inventory Adjust,Invoice,Invoice Payment,Pay Bill,POS,Purchase Bill Payment,Receive Payment,Retained Earning 
                                    ws_gld_data.Cells(gld_inventory_id,17) = 0 'debit
                                    ws_gld_data.Cells(gld_inventory_id,18) = totalCost 'credit
                                    ws_gld_data.Cells(gld_inventory_id,19) = "ICS: Inventory for INV # " & invoiceCode & " Product # " & productCode & " " &  productName 'memo
                                    ws_gld_data.Cells(gld_inventory_id,20) = customerId 'customer_id
                                    ws_gld_data.Cells(gld_inventory_id,24) = 1 'class_id
                                    ws_gld_data.Cells(gld_inventory_id,25) = 1 'is_active
                                    ws_gld_data.Cells(gld_inventory_id,26) = invoiceWeek 'week
                                    ws_gld_data.Cells(gld_inventory_id,27) = invoiceMonth 'month
                                    ws_gld_data.Cells(gld_inventory_id,28) = invoiceYear 'year
                                    gld_inventory_id = gld_inventory_id + 1
                                    
                                    gld_cogs_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                    ws_gld_data.Cells(gld_cogs_id,1)  = gld_cogs_id 'id
                                    ws_gld_data.Cells(gld_cogs_id,2)  = gl_id 'general_ledger_id
                                    ws_gld_data.Cells(gld_cogs_id,3)  = 11 'chart_account_id application.VLOOKUP("Sales Income",chart_account_default_list2,3,FALSE)
                                    ws_gld_data.Cells(gld_cogs_id,4)  = 1 'company_id
                                    ws_gld_data.Cells(gld_cogs_id,5)  = branchId 'branch id
                                    ws_gld_data.Cells(gld_cogs_id,6)  = locationGroupId 'location group id
                                    ws_gld_data.Cells(gld_cogs_id,7)  = locationId 'location id
                                    ws_gld_data.Cells(gld_cogs_id,8)  = productId 'product_id
                                    ws_gld_data.Cells(gld_cogs_id,11) = invValId 'inventory_valuation_id
                                    ws_gld_data.Cells(gld_cogs_id,12) = 1 'inventory_valuation_is_debit
                                    ws_gld_data.Cells(gld_cogs_id,14) = invDetailId 'sales_invoice_detail_id
                                    ws_gld_data.Cells(gld_cogs_id,16) = "Invoice" 'type
                                    ws_gld_data.Cells(gld_cogs_id,17) = totalCost 'debit
                                    ws_gld_data.Cells(gld_cogs_id,18) = 0 'credit
                                    ws_gld_data.Cells(gld_cogs_id,19) = "ICS: COGS for INV # " & invoiceCode & " Product # " & productCode & " "  & productName 'memo
                                    ws_gld_data.Cells(gld_cogs_id,20) = customerId 'customer_id
                                    ws_gld_data.Cells(gld_cogs_id,24) = 1 'class_id
                                    ws_gld_data.Cells(gld_cogs_id,25) = 1 'is_active
                                    ws_gld_data.Cells(gld_cogs_id,26) = invoiceWeek 'week
                                    ws_gld_data.Cells(gld_cogs_id,27) = invoiceMonth 'month
                                    ws_gld_data.Cells(gld_cogs_id,28) = invoiceYear 'year
                                    gld_cogs_id = gld_cogs_id + 1
                                End If
                                invValId = invValId + 1

                                'General Ledger Detail (Product Income)
                                If (totalPriceBeforDis > 0) Then
                                    gld_sid_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                    ws_gld_data.Cells(gld_sid_id,1)  = gld_sid_id 'id
                                    ws_gld_data.Cells(gld_sid_id,2)  = gl_id 'general_ledger_id
                                    ws_gld_data.Cells(gld_sid_id,3)  = coaSaleIncome 'chart_account_id application.VLOOKUP("Sales Income",chart_account_default_list2,3,FALSE)
                                    ws_gld_data.Cells(gld_sid_id,4)  = 1 'company_id
                                    ws_gld_data.Cells(gld_sid_id,5)  = branchId 'branch id
                                    ws_gld_data.Cells(gld_sid_id,6)  = locationGroupId 'location group id
                                    ws_gld_data.Cells(gld_sid_id,7)  = locationId 'location id
                                    ws_gld_data.Cells(gld_sid_id,8)  = productId 'product_id
                                    ws_gld_data.Cells(gld_sid_id,14) = invDetailId 'sales_invoice_detail_id
                                    ws_gld_data.Cells(gld_sid_id,16) = "Invoice" 'type
                                    ws_gld_data.Cells(gld_sid_id,17) = 0 'debit
                                    ws_gld_data.Cells(gld_sid_id,18) = totalPriceBeforDis 'credit
                                    ws_gld_data.Cells(gld_sid_id,19) = "ICS: INV # " & invoiceCode & " Product # " & productCode & " "  & productName 'memo
                                    ws_gld_data.Cells(gld_sid_id,20) = customerId 'customer_id
                                    ws_gld_data.Cells(gld_sid_id,24) = 1 'class_id
                                    ws_gld_data.Cells(gld_sid_id,25) = 1 'is_active
                                    ws_gld_data.Cells(gld_sid_id,26) = invoiceWeek 'week
                                    ws_gld_data.Cells(gld_sid_id,27) = invoiceMonth 'month
                                    ws_gld_data.Cells(gld_sid_id,28) = invoiceYear 'year
                                    gld_sid_id = gld_sid_id + 1
                                End If
                                
                                 'General Ledger Detail (Product Discount)
                                If (totalDisItem > 0) Then
                                    'insert general ledger detail Accounts Receivable
                                    gld_discount_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                    ws_gld_data.Cells(gld_discount_id,1)  = gld_discount_id 'id
                                    ws_gld_data.Cells(gld_discount_id,2)  = gl_id 'general_ledger_id
                                    ws_gld_data.Cells(gld_discount_id,3)  = coaSaleDiscount 'chart_account_id
                                    ws_gld_data.Cells(gld_discount_id,4)  = 1 'company_id
                                    ws_gld_data.Cells(gld_discount_id,5)  = branchId 'branch id
                                    ws_gld_data.Cells(gld_discount_id,6)  = locationGroupId 'location group id
                                    ws_gld_data.Cells(gld_discount_id,7)  = locationId 'location id
                                    ws_gld_data.Cells(gld_discount_id,14) = invDetailId 'sales_invoice_detail_id
                                    ws_gld_data.Cells(gld_discount_id,16) = "Invoice" 'type
                                    ws_gld_data.Cells(gld_discount_id,17) = totalDisItem 'debit
                                    ws_gld_data.Cells(gld_discount_id,18) = 0 'credit
                                    ws_gld_data.Cells(gld_discount_id,19) = "ICS: INV # " & invoiceCode & " Product # " & productCode & " " & productName & " Discount" 'memo
                                    ws_gld_data.Cells(gld_discount_id,20) = customerId 'customer_id
                                    ws_gld_data.Cells(gld_discount_id,24) = 1 'class_id
                                    ws_gld_data.Cells(gld_discount_id,25) = 1 'is_active
                                    ws_gld_data.Cells(gld_discount_id,26) = invoiceWeek 'week
                                    ws_gld_data.Cells(gld_discount_id,27) = invoiceMonth 'month
                                    ws_gld_data.Cells(gld_discount_id,28) = invoiceYear 'year
                                    gld_discount_id = gld_discount_id + 1
                                End If

                                invDetailId = invDetailId + 1

                                'insert inventory
                                Set inventory  = ws_inv.Range("C8:AD8")
                                inventoryId    = ws_inv.Cells(Rows.Count, "C").End(xlUp).Row - 6

                                inventory.Cells(inventoryId, 1)   = inventoryId 'id
                                inventory.Cells(inventoryId, 2)   = productId 'product id
                                inventory.Cells(inventoryId, 3)   = productCode 'product code
                                inventory.Cells(inventoryId, 4)   = productName 'product name
                                inventory.Cells(inventoryId, 5)   = invoiceDate 'inventory date
                                inventory.Cells(inventoryId, 6)   = locationId 'location id
                                inventory.Cells(inventoryId, 7)   = locationGroupId 'location group id
                                inventory.Cells(inventoryId, 8)   = "Sale" 'inventory type (Inv Adj,Purchase,Sale,Void Sale...)
                                inventory.Cells(inventoryId, 9)   = customerId 'customer id
                                inventory.Cells(inventoryId, 12)  = invId 'Sale Invoice Id
                                inventory.Cells(inventoryId, 19)  = qtyOrderSmall * -1 'Qty
                                inventory.Cells(inventoryId, 20)  = unitCostByUom 'Unit Cost 
                                inventory.Cells(inventoryId, 21)  = unitPrice 'Unit Price
                                inventory.Cells(inventoryId, 22)  = 1 'status
                                inventory.Cells(inventoryId, 23)  = lotsNumber 'Lots Number, 
                                If (expiredDate > 0) Then 
                                inventory.Cells(inventoryId, 24)  = expiredDate 'Expired Date
                                Else
                                inventory.Cells(inventoryId, 24)  = "" 'Expired Date
                                End If
                                inventory.Cells(inventoryId, 25)  = createdDate 'created date
                                inventory.Cells(inventoryId, 26)  = invoiceWeek 'week
                                inventory.Cells(inventoryId, 27)  = invoiceMonth 'month
                                inventory.Cells(inventoryId, 28)  = invoiceYear 'year
                                inventoryId = inventoryId + 1

                                'Insert / Update 1_group_totals
                                Set invGroupTotal        = ws_group_total.Range("C8:S8")
                                invGroupTotalId          = ws_group_total.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                LastRowInvGroupTotal     = ws_group_total.Cells(Rows.Count, "C").End(xlUp).Row

                                ' checkExistItemInvGroupTotal=0 ,LastRowInvGroupTotal=10
                                ' Insert group total =>productId=1 , invGroupTotalId=4
                                ' checkExistItemInvGroupTotal=0 ,LastRowInvGroupTotal=11
                                ' Insert group total =>productId=2 , invGroupTotalId=5
                                ' checkExistItemInvGroupTotal=0 ,LastRowInvGroupTotal=12
                                ' Insert group total =>productId=3 , invGroupTotalId=6

                                ' debug.Print "checkExistItemInvGroupTotal=" & checkExistItemInvGroupTotal & " ,LastRowInvGroupTotal=" & LastRowInvGroupTotal
                                If checkExistItemInvGroupTotal > 0 Then
                                    qtyInvAdj    = 0
                                    qtySale      = 0
                                    qtySaleFree  = 0
                                    qtyPos       = 0
                                    qtyPosFree   = 0
                                    qtyPb        = 0
                                    qtyPr        = 0
                                    qtySr        = 0
                                    qtySrFree    = 0
                                    qtyToIn      = 0
                                    qtyToOut     = 0
                                    For i = 8 To LastRowInvGroupTotal
                                        'update inventory total locationGroupId locationId
                                        ' debug.Print "product id=" & ws_group_total.Cells(i,3).Value
                                        ' Debug.Print "Product=>" & ws_group_total.Cells(i,3).Value * 1 & "=" & productId * 1  & ",locationGroupId=>" &   ws_group_total.Cells(i,6).Value & "=" &  locationGroupId   & ",locationId=>" &  ws_group_total.Cells(i,7).Value & "-" & locationId
                                        If ws_group_total.Cells(i,3).Value * 1 = productId * 1 And ws_group_total.Cells(i,6).Value = locationGroupId And ws_group_total.Cells(i,7).Value = locationId And ws_group_total.Cells(i,8).Value = lotsNumber  And ws_group_total.Cells(i,9).Value = expiredDate Then
                                            ' Debug.Print "Update group total"
                                            qtyInvAdj    = ws_group_total.Cells(i, 11) * 1
                                            qtySale      = ws_group_total.Cells(i, 12) * 1
                                            qtyPos       = ws_group_total.Cells(i, 13) * 1
                                            qtyPb        = ws_group_total.Cells(i, 14) * 1
                                            qtyPr        = ws_group_total.Cells(i, 15) * 1
                                            qtySr        = ws_group_total.Cells(i, 16) * 1
                                            qtyToIn      = ws_group_total.Cells(i, 17) * 1
                                            qtyToOut     = ws_group_total.Cells(i, 18) * 1

                                            stockIn      = qtyInvAdj + qtyPb + qtySr + qtyToIn
                                            stockOut     = qtyOrderSmall + (qtySale + qtyPos + qtyPr + qtyToOut) '(qty + qtyFree) = New Qty Sales
                                            totalQty     = stockIn - stockOut

                                            ws_group_total.Cells(i,10)  = totalQty 'total qty
                                            ws_group_total.Cells(i,12)  = qtySale + qtyOrderSmall 'total qty sale
                                            Exit For
                                        End If
                                    Next i
                                Else
                                    'insert New record inventory group total invGroupTotal
                                    ' Debug.Print "Insert group total =>productId=" & productId & " , invGroupTotalId=" & invGroupTotalId
                                    invGroupTotal.Cells(invGroupTotalId,1)   = productId 'product id 
                                    invGroupTotal.Cells(invGroupTotalId,2)   = productCode 'product code
                                    invGroupTotal.Cells(invGroupTotalId,3)   = productName 'product name
                                    invGroupTotal.Cells(invGroupTotalId,4)   = locationGroupId 'location group id
                                    invGroupTotal.Cells(invGroupTotalId,5)   = locationId 'location id
                                    If (lotsNumber<> "") Then 
                                    invGroupTotal.Cells(invGroupTotalId,6)   = lotsNumber 'lote number
                                    Else
                                    invGroupTotal.Cells(invGroupTotalId,6)   = "" 'lote number
                                    End If
                                    If (expiredDate >0) Then 
                                    invGroupTotal.Cells(invGroupTotalId,7)   = expiredDate 'expired date
                                    Else
                                    invGroupTotal.Cells(invGroupTotalId,7)   = "" 'expired date
                                    End If
                                    invGroupTotal.Cells(invGroupTotalId,8)   = qtyOrderSmall * -1'total qty ending
                                    invGroupTotal.Cells(invGroupTotalId,9)   = 0 'total qty adjustment
                                    invGroupTotal.Cells(invGroupTotalId,10)  = (qty + qtyFree) * (smallValUom / conversion) 'total qty sale
                                    invGroupTotal.Cells(invGroupTotalId,11)  = 0 'Total Pos
                                    invGroupTotal.Cells(invGroupTotalId,12)  = 0 'Total PB
                                    invGroupTotal.Cells(invGroupTotalId,13)  = 0 'Total PR
                                    invGroupTotal.Cells(invGroupTotalId,14)  = 0 'Total SR
                                    invGroupTotal.Cells(invGroupTotalId,15)  = 0 'Total TO In
                                    invGroupTotal.Cells(invGroupTotalId,16)  = 0 'Total TO Out
                                    invGroupTotal.Cells(invGroupTotalId,17)  = 0 'Total Order
                                    invGroupTotalId = invGroupTotalId + 1
                                End If

                                'insert / update 1_group_total_detail
                                Set invTotalGroupDetail  = ws_group_total_detail.Range("C8:Q8")
                                invTotalGroupDetailId    = ws_group_total_detail.Cells(Rows.Count, "C").End(xlUp).Row - 6

                                lastInvTotalGroupDetailId = ws_group_total_detail.Cells(Rows.Count, "C").End(xlUp).Row
                                ' debug.Print "checkExistItemInvGroupTotalDetail=" & checkExistItemInvGroupTotalDetail & ",lastInvTotalGroupDetailId" & lastInvTotalGroupDetailId
                                If checkExistItemInvGroupTotalDetail > 0  Then
                                    For j = 8 To lastInvTotalGroupDetailId
                                        If (ws_group_total_detail.Cells(j,3).Value * 1 = productId * 1 And ws_group_total_detail.Cells(j,6).Value * 1 = locationGroupId And ws_group_total_detail.Cells(j,7).Value * 1 = locationId And ws_group_total_detail.Cells(j,8).Value = invoiceDate) Then
                                            'update  inventory total detail
                                            'Debug.Print "Update group total detail"
                                            ws_group_total_detail.Cells(j,10) = ws_group_total_detail.Cells(j,10) * 1 + qtyOrderSmall 'total qty sale 
                                            Exit For
                                        End If
                                    Next j
                                Else
                                    'insert New record inventory total detail
                                    ' Debug.Print "Insert group total detail"
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,1)   = productId 'product id
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,2)   = productCode 'product code
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,3)   = productName 'product name
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,4)   = locationGroupId 'location id
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,5)   = locationId 'location id
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,6)   = invoiceDate 'date
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,7)   = 0 'total adjustment
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,8)   = qtyOrderSmall 'total sale invoice
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,9)   = 0 'Total Pos
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,10)  = 0 'Total Purchase Bill
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,11)  = 0 'Total Purchase Return
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,12)  = 0 'Total Sales Return
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,13)  = 0 'Total Transfer Order In
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,14)  = 0 'Total Transfer Order Out
                                    invTotalGroupDetail.Cells(invTotalGroupDetailId,15)  = 0 'Total Order
                                    invTotalGroupDetailId = invTotalGroupDetailId + 1
                                End If

                                'Insert/update 1_inventory_totals
                                Set invTotal         = ws_invTotal.Range("C8:T8")
                                invTotalId           = ws_invTotal.Cells(Rows.Count, "C").End(xlUp).Row - 6

                                LastRowInvTotal = ws_invTotal.Cells(Rows.Count, "C").End(xlUp).Row
                                ' Debug.Print "checkExistItemInvTotal=" & checkExistItemInvTotal
                                If checkExistItemInvTotal > 0 Then
                                    qtyInvAdj    = 0
                                    qtySale      = 0
                                    qtySaleFree  = 0
                                    qtyPos       = 0
                                    qtyPosFree   = 0
                                    qtyPb        = 0
                                    qtyPr        = 0
                                    qtySr        = 0
                                    qtySrFree    = 0
                                    qtyToIn      = 0
                                    qtyToOut     = 0
                                    For i = 8 To LastRowInvTotal
                                        'update inventory total
                                        If ws_invTotal.Cells(i,3).Value * 1 = productId * 1 And ws_invTotal.Cells(i,6).Value = lotsNumber  And ws_invTotal.Cells(i,7).Value = expiredDate Then
                                            ' Debug.Print "Update inventory total"
                                            qtyInvAdj    = ws_invTotal.Cells(i, 9) * 1
                                            qtySale      = ws_invTotal.Cells(i, 10) * 1
                                            qtySaleFree  = ws_invTotal.Cells(i, 11) * 1
                                            qtyPos       = ws_invTotal.Cells(i, 12) * 1
                                            qtyPosFree   = ws_invTotal.Cells(i, 13) * 1
                                            qtyPb        = ws_invTotal.Cells(i, 14) * 1
                                            qtyPr        = ws_invTotal.Cells(i, 15) * 1
                                            qtySr        = ws_invTotal.Cells(i, 16) * 1
                                            qtySrFree    = ws_invTotal.Cells(i, 17) * 1
                                            qtyToIn      = ws_invTotal.Cells(i, 16) * 1
                                            qtyToOut     = ws_invTotal.Cells(i, 17) * 1

                                            stockIn      = qtyInvAdj + qtyPb + qtySr + qtySrFree + qtyToIn 
                                            stockOut     = qtyOrderSmall + (qtySale + qtySaleFree + qtyPos + qtyPosFree + qtyPr + qtyToOut) '(qty + qtyFree) = New Qty Sales
                                            totalQty     = stockIn - stockOut

                                            ws_invTotal.Cells(i,8)  = totalQty 'total qty
                                            ws_invTotal.Cells(i,10) = qtySale + qty * (smallValUom / conversion) 'total qty sale
                                            ws_invTotal.Cells(i,11) = qtySaleFree  + qtyFree * (smallValUom / conversion) 'total qty sale free
                                            Exit For
                                        End If
                                    Next i
                                Else
                                    'insert New record inventory total
                                    ' Debug.Print "Insert inventory total"
                                    invTotal.Cells(invTotalId,1)   = productId 'product id 
                                    invTotal.Cells(invTotalId,2)   = productCode 'product code
                                    invTotal.Cells(invTotalId,3)   = productName 'product name
                                    invTotal.Cells(invTotalId,4)   = lotsNumber 'lot number
                                    If (expiredDate > 0) Then 
                                    invTotal.Cells(invTotalId,5)   = expiredDate 'expired date
                                    Else
                                    invTotal.Cells(invTotalId,5)   = "" 'expired date
                                    End If
                                    invTotal.Cells(invTotalId,6)   = qtyOrderSmall * -1'total qty ending
                                    invTotal.Cells(invTotalId,7)   = 0 'total qty adjustment
                                    invTotal.Cells(invTotalId,8)   = qty * (smallValUom / conversion) 'total qty sale
                                    invTotal.Cells(invTotalId,9)   = qtyFree * (smallValUom / conversion) 'total qty free sale
                                    invTotal.Cells(invTotalId,10)  = 0 'Total Pos
                                    invTotal.Cells(invTotalId,11)  = 0 'Total Pos Free
                                    invTotal.Cells(invTotalId,12)  = 0 'Total PB
                                    invTotal.Cells(invTotalId,13)  = 0 'Total PR
                                    invTotal.Cells(invTotalId,14)  = 0 'Total SR
                                    invTotal.Cells(invTotalId,15)  = 0 'Total SR Free
                                    invTotal.Cells(invTotalId,16)  = 0 'Total TO In
                                    invTotal.Cells(invTotalId,17)  = 0 'Total TO Out
                                    invTotal.Cells(invTotalId,18)  = 0 'Total Order
                                    invTotalId = invTotalId + 1
                                End If

                                'insert/update inventory total detail
                                Set invTotalDetail   = ws_inv_total_detail.Range("C8:R8")
                                invTotalDetailId     = ws_inv_total_detail.Cells(Rows.Count, "C").End(xlUp).Row - 6

                                lastRowInvTotalDetail = ws_inv_total_detail.Cells(Rows.Count, "C").End(xlUp).Row
                                ' debug.Print "checkExistItemInvTotalDetail=" & checkExistItemInvTotalDetail & ",lastRowInvTotalDetail=" & lastRowInvTotalDetail
                                If checkExistItemInvTotalDetail > 0  Then
                                    For j = 8 To lastRowInvTotalDetail
                                        If (ws_inv_total_detail.Cells(j,3).Value * 1 = productId * 1 And ws_inv_total_detail.Cells(j,7).Value = lotsNumber And ws_inv_total_detail.Cells(j,8).Value  = expiredDate And ws_inv_total_detail.Cells(j,9).Value = invoiceDate) Then
                                            'update  inventory total detail
                                            ' Debug.Print "Update inventory total detail"
                                            ws_inv_total_detail.Cells(j,11) = ws_inv_total_detail.Cells(j,11) * 1 + qtyOrderSmall 'total qty sale 
                                            Exit For
                                        End If
                                    Next j
                                Else
                                    'insert New record inventory total detail
                                    ' Debug.Print "Insert inventory total detail"
                                    invTotalDetail.Cells(invTotalDetailId,1)   = productId 'product id
                                    invTotalDetail.Cells(invTotalDetailId,2)   = productCode 'product code
                                    invTotalDetail.Cells(invTotalDetailId,3)   = productName 'product name
                                    invTotalDetail.Cells(invTotalDetailId,4)   = locationId 'location id
                                    If (lotsNumber<> "") Then 
                                    invTotalDetail.Cells(invTotalDetailId,5)   = lotsNumber 'lots number
                                    Else
                                    invTotalDetail.Cells(invTotalDetailId,5)   = "" 'lots number
                                    End If
                                    If (expiredDate >0) Then 
                                    invTotalDetail.Cells(invTotalDetailId,6)   = expiredDate 'expired date
                                    Else
                                    invTotalDetail.Cells(invTotalDetailId,6)   = "" 'expired date
                                    End If
                                    invTotalDetail.Cells(invTotalDetailId,7)   = invoiceDate 'date
                                    invTotalDetail.Cells(invTotalDetailId,8)   = 0 'total adjustment
                                    invTotalDetail.Cells(invTotalDetailId,9)   = qtyOrderSmall 'total sale invoice
                                    invTotalDetail.Cells(invTotalDetailId,10)  = 0 'Total Pos
                                    invTotalDetail.Cells(invTotalDetailId,11)  = 0 'Total Purchase Bill
                                    invTotalDetail.Cells(invTotalDetailId,12)  = 0 'Total Purchase Return
                                    invTotalDetail.Cells(invTotalDetailId,13)  = 0 'Total Sales Return
                                    invTotalDetail.Cells(invTotalDetailId,14)  = 0 'Total Transfer Order In
                                    invTotalDetail.Cells(invTotalDetailId,15)  = 0 'Total Transfer Order Out
                                    invTotalDetail.Cells(invTotalDetailId,16)  = 0 'Total Order
                                    invTotalDetailId = invTotalDetailId + 1
                                End If
                            End If
                        Next a
                        '*******End save date To sales by item data*******
                        access = 1
                    End If

                    Call resetCellValue(activeWorkbook.Sheets("AddInvoice"))
                    ws_invoice_data.Select
                    invId = ws_invoice_data.Cells(Rows.Count, "C").End(xlUp).Row
                    ws_invoice_data.Range("C" & invId & ":" & "AI" & invId).Select 'Select New insert invoice row

                    wbDatabase.Save
                    OnEnd
                    MsgBox "Invoice " & invoiceCode & " saved successful."
                    Exit Sub
                End If
            End If
        End If
    Else
        Exit Sub
    End If
End Sub

Private Function setStatusColor(rngInvoice As Range)
    Dim iset As IconSetCondition

    ' rngInvoice.Select
    rngInvoice.FormatConditions.Delete
    Set iset = rngInvoice.FormatConditions.AddIconSetCondition

    'Select the traffic lights iconset
    With iset
        .IconSet = ActiveWorkbook.IconSets(xl4TrafficLights)
        .ReverseOrder = False
        .ShowIconOnly = True
    End With

    ' xlIconBlackCircleWithBorder , xlIconGrayCircle , xlIconGreenCircle , xlIconRedCircleWithBorder , xlIconPinkCircle , xlIconYellowCircle , xlIconGreenCheckSymbol , xlIconRedCrossSymbol , xlIconYellowExclamationSymbol , xlIconWhiteCircleAllWhiteQuarters
    ' Sales Invoice Status=> -1 = Edit, 1  = issue, 2  = fulfilled ,3  = partial

    With iset.IconCriteria(2)
        .Icon = xlIconBlackCircleWithBorder
        .Type = xlConditionValueFormula
        .Operator = xlGreaterEqual
        .Value = "=-1" 'edit
    End With

    With iset.IconCriteria(2)
        .Icon = xlIconRedCircleWithBorder
        .Type = xlConditionValueFormula
        .Operator = xlGreaterEqual
        .Value = "=1" 'issue
    End With

    With iset.IconCriteria(3)
        .Icon = xlIconGreenCheckSymbol
        .Type = xlConditionValueFormula
        .Operator = xlGreaterEqual
        .Value = "=2" 'fulfilled
    End With

    With iset.IconCriteria(4)
        .Icon = xlIconYellowCircle
        .Type = xlConditionValueFormula
        .Operator = xlGreaterEqual
        .Value = "=3" 'partial
    End With
End Function

Private Function resetCellValue(ws_add_inv AS Worksheet)
    With ws_add_inv
        .Range("invoice_column_product").Value = ""
        .Range("invoice_qty").Value = ""
        .Range("invoice_qty_free").Value =""
        .Range("invoice_column_uom").Value = ""
        .Range("invoice_column_new_unit_price").Value = ""
        .Range("invoice_discount_by_item").Value = ""
        .Range("invoice_note").Value = ""
        .Range("invoice_customer").value = .Range("C25").value
        .Range("invoice_column_expired_date").Value = ""
        .Range("invoice_column_lote_number").Value = ""
        .Range("invoice_column_remark").Value = ""
    End With
End Function

Private Sub getExchangeRateNBC()
    ActiveSheet.Range("invoice_exchange_rate").Value = Sheets("NBC_Exchange_Rate").Range("ExchangeRateToday").Value
End Sub

Private Sub TodayPickdate()
    ActiveSheet.Range("invoice_date").Value = Format(Now(), "yyyy/mm/dd")
End Sub

Private Sub copyInvoiceForm()
    Sheets("AddInvoice").Range("invoice_copy_area").Copy
End Sub

' Private Function updateStatusSalesInvoice(invoiceNo)
'     Dim invId, lastCol,invIdSd, m, n As Long
'     Dim ws,sd As Worksheet

'     Set ws = Sheets("SaleInvoiceData")
'     Set sd = Sheets("SalesByItemData")

'     'Sales Invoice Status=> -1 = Edit, 1  = issue, 2  = fulfilled ,3  = partial
'     invId = ws.Cells(Rows.Count, "C").End(xlUp).Row 
'     lastCol = ws.Cells(6, Columns.Count).End(XlToLeft).column 'start from row 6 column B
'     For m = 6 To invId
'         If ws.Cells(m, 3).Value = invoiceNo*1 Then
'             ws.Cells(m, 3).Value = invoiceNo*(-1)
'             'Update style cell value
'             With ws.Range("B" & m).Font
'                 .Bold = True
'                 .ColorIndex = 3
'                 .Name = "Roboto"
'                 .Size = 13
'             End With
'             ws.Cells(m, 17).Value = -1  'Update status sales invoice
'         End If
'     Next m

'     invIdSd = sd.Cells(Rows.Count, "A").End(xlUp).Row 
'     For n = 3 To invIdSd
'         If sd.Cells(n, 3).Value = invoiceNo*1 Then
'             'Update style cell value
'             With sd.Range("A" & n).Font
'                 .Bold = True
'                 .ColorIndex = 3
'                 .Name = "Roboto"
'                 .Size = 13
'             End With
'             sd.Cells(n, 3).Value = invoiceNo*(-1)  'Update status sales invoice detail
'         End If
'     Next n
' End Function
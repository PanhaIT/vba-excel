Attribute VB_Name = "add_purchase_bill"

Public  Sub Worksheet_SelectionChange(ByVal Target As Range)
    cells.EntireColumn.AutoFit
End Sub

Private Sub getCurrentDate()
    ActiveSheet.Range("pb_date").Value = Format(Now(), "dd/mm/yyyy")
End Sub

Private Sub getPbNBCExchangeRate()
    ActiveSheet.Range("pb_exchange_rate").Value = Sheets("NBC_Exchange_Rate").Range("ExchangeRateToday").Value
End Sub

Private Sub SavePurchaseBill()
    OnStart
    Dim ws_gld, ws_gl,ws_product,ws_pro_cost_his As Worksheet
    Dim defaultCost,newUnitCost,unitCostByUoM,discountAmount,discountPercent,unitPrice, totalCost, totalDeposit,totalDiscount,totalAmount,grandTotalAmount,unitCost,totalDisItem,disItem,totalVat As Double
    Dim statusNote,vendorId,pbSheet,branchId,locationId,priceTypeId,discountTypeId As string
    Dim vendorName,pbNote,pbWeek,pbMonth,pbYear,expenseTypeId,expenseType,branchName,msgSave As String
    Dim conversion, qtyOrder,qtyOrderSmall,inv_valuation_id,statusPB, pbNo, indexColor, pb_id, pb_detail_id, inventory_id, inv_total_detail_id, inventory_total_id,itemType As Integer

    Dim qtyInvAdj,qtySale,qtySaleFree, qtyPb, qtyPr, qtySr, qtySrFree, qtyToIn, qtyToOut,stockIn,stockOut,totalQty,coaPurchaseVat,gld_sid_id,gld_cogs_id,gld_inventory_id As Integer
    Dim checkExistProductInvTotal,checkExistProductInvTotalDetail, checkExistItemInvGroupTotal,checkExistItemInvGroupTotalDetail,stockByDateInvGroupTotalDetail,stockAvailableInvGroupTotal ,paymentTermId, qtyStock, productId, qty, qtyFree,smallValUoM, vatPercent,vatSettingId,vatCal,chartAccIncome,gld_sales_dis_id As Integer
    Dim purchaseRequestNo, refinvoiceCode ,productCode, productName, productUoM,pbCode As String

    Dim invGroupTotal,invTotalGroupDetail, item_row, range_pb, pb_data_r, range_inventory_total,ws_gld_data,ws_gl_data, range_inventory_valuation,productRange,proCostHisRange As Range
    Dim i, j, a, b, LastRowInvTotal, lastRowInvTotalDetail  As Long
    Dim LastRowInvGroupTotal,lastInvTotalGroupDetail, invGroupTotalId,invTotalGroupDetailId, gld_discount_id, coaPurchaseDiscount, netDays,last_product_id,last_pro_cost_his As Integer
    Dim ws_pb, ws_pb_total, ws_pb_total_detail,ws_pb_valuaction,ws_group_total_detail,ws_group_total As Worksheet
    Dim totalAmountItem,costOrder As Double
    Dim coaDefaultRange As Range
    Dim productUomId As Integer

    Dim pbDate, createdDate,pbRefInvoiceDate,pbNetDays As Date
    Dim answer As VbMsgBoxResult

    Dim activeWorkbook, wbDatabase As Workbook
    Dim checkExistDb As Boolean
    Dim ws_file_setting,ws_purchase_bill As Worksheet
    Dim MyFSO As New FileSystemObject
    Dim databasePath, databaseFile, remark As String

    Set activeWorkbook = Workbooks("index.xlsm")

    msgSave           = "Are you want to save this purchase bill?"
    answer            = MsgBoxW(msgSave, vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Add Sales pb")
    pbNo              = ActiveSheet.Range("pb_no").Value
    pbCode            = ActiveSheet.Range("pb_code").Value
    vendorId          = ActiveSheet.Range("pb_vendor_id").Value
    'pbDate           = Format(ActiveSheet.Range("pb_date").Value, "dd/mm/yyyy")
    branchId          = ActiveSheet.Range("pb_branch_id").Value

    If (ActiveSheet.Range("pb_branch").Value <> "") Then 
    branchName        = MID(ActiveSheet.Range("pb_branch").Value,5,100)
    Else
        MsgBox "Branch is requied, try again."
        Exit Sub
    End If
    
    locationGroupId   = 1

    If (ActiveSheet.Range("pb_location_id").Value <> "") Then 
    locationId        = MID(ActiveSheet.Range("pb_location_id").Value,1,2) * 1
    Else
        MsgBox "Location is requied, try again."
        Exit Sub
    End If

    purchaseRequestNo = ActiveSheet.Range("pb_purchase_request_no").Value
    refinvoiceCode    = ActiveSheet.Range("pb_ref_invoice").Value
    pbDate            = ActiveSheet.Range("pb_date").Value

    If (ActiveSheet.Range("pb_ref_invoice_date").Value <> "") Then 
        pbRefInvoiceDate = ActiveSheet.Range("pb_ref_invoice_date").Value
    Else
        pbRefInvoiceDate = 0
    End If

    discountTypeId    = ActiveSheet.Range("pb_discount_id").Value
    discountAmount    = ActiveSheet.Range("pb_discount_amount").Value
    discountPercent   = ActiveSheet.Range("pb_discount_percent").Value

    If (ActiveSheet.Range("pb_vendor").Value <> "") Then 
        vendorName        = MID(ActiveSheet.Range("pb_vendor").Value,5,100)
    Else
        MsgBox "Vendor is requied, try again."
        Exit Sub
    End If

    pbNote            = ActiveSheet.Range("pb_note").Value
    paymentTermId     = ActiveSheet.Range("pb_payment_term_id").Value
    currencyId        = ActiveSheet.Range("pb_currency_id").Value

    If (ActiveSheet.Range("pb_date").Value >0) Then 
        pbMonth       = month(ActiveSheet.Range("pb_date").Value)
        pbYear        = year(ActiveSheet.Range("pb_date").Value)
    Else
        MsgBox "Purchase date is requied, try again."
        Exit Sub
    End If

    pbWeek            = ActiveSheet.Range("pb_week").Value
    pbNetDays         = ActiveSheet.Range("pb_net_days").Value
    vatPercent        = ActiveSheet.Range("pb_vat_percent").Value

    If (ActiveSheet.Range("pb_vat").Value <> "") Then 
        vatSettingId   = MID(ActiveSheet.Range("pb_vat").Value,1,2)*1
    Else
        vatSettingId  = 1
    End If

    vatCal            = ActiveSheet.Range("pb_vat_cal_id").Value
    totalVat          = ActiveSheet.Range("pb_total_vat").Value

    totalDeposit      = ActiveSheet.Range("pb_deposit").Value
    totalDiscount     = ActiveSheet.Range("pb_discount").Value
    totalAmount       = ActiveSheet.Range("pb_sub_total_amount").Value
    statusPB          = ActiveSheet.Range("pb_status").Value
    grandTotalAmount  = ActiveSheet.Range("pb_total_amount").Value

    Set coaDefaultRange = Sheets("ChartAccountDefault").Range("chart_account_default_list2")

    coaPurchaseDiscount = Application.VLookup("Purchase Discount", coaDefaultRange, 3, False)
    coaInventory        = Application.VLookup("Inventory Asset Account", coaDefaultRange, 3, False)
    coaPayable          = Application.VLookup("Accounts Payable", coaDefaultRange, 3, False)
    coaPurchaseVat      = Application.VLookup("Purchase VAT", coaDefaultRange, 3, False)
    
    If (ActiveSheet.Range("pb_exchange_rate").Value >0) Then 
        exchangeRate  = ActiveSheet.Range("pb_exchange_rate").Value
    Else
        exchangeRate  = 4000
    End If
   
    totalDeposit = 0
    totalBalance = totalAmount + totalVat - (totalDiscount + totalDeposit)
    
    If answer = vbYes And checkDatabaseFile = TRUE Then
        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath = ws_file_setting.Range("database_path").Value
        databaseFile = databasePath & "database_kscm.xlsm"

        If (IsWorkbookOpen(CStr(databaseFile)) = FALSE) Then
            Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        End If

        Run resetAllFilter()

        Set ws_product          = activeWorkbook.Sheets("ProductList")
        Set rngProductList      = ws_product.Range("C8:AA" & ws_product.Cells(Rows.Count,"C").End(xlUp).Row)
        Set ws_pro_cost_his     = activeWorkbook.Sheets("ProductCostHistory")

        Set ws_purchase_bill    = wbDatabase.Sheets("PurchaseBill")
        Set ws_pb_detail        = wbDatabase.Sheets("PurchaseBillDetail")

        Set ws_gl               = wbDatabase.sheets("GeneralLedger")
        Set ws_gld              = wbDatabase.Sheets("GeneralLedgerDetail")

        Set ws_pb               = wbDatabase.Sheets("inventories")
        Set ws_pb_total         = wbDatabase.Sheets(locationId & "_inventory_totals")
        Set ws_pb_total_detail  = wbDatabase.Sheets(locationId & "_inventory_total_details")

        Set ws_group_total        = wbDatabase.Sheets(locationGroupId & "_group_totals")
        Set ws_group_total_detail = wbDatabase.Sheets(locationGroupId & "_group_total_details")

        Set ws_pb_valuaction    = wbDatabase.Sheets("InventoryValuation")

        createdDate             = Format(Now(), "yyyy/mm/dd hh:mm:ss")
        
        'Sales pb Status=> -1 = Edit, 1  = issue, 2  = fulfilled ,3  = partial
        ' statusPB      = -1
        ' indexColor    = 1
        ' statusNote    = "Edit"
        If totalBalance = 0 Then
            statusPB    = 2
            indexColor  = 43
            statusNote  = "fulfilled"
        Elseif totalBalance > 0 Then
            If totalDeposit > 0 AND  totalDeposit < (totalAmount + totalVat - totalDiscount) Then 
                statusPB   = 3
                indexColor = 44
                statusNote = "Partial"
            ElseIf totalDeposit = 0 Then
                statusPB   = 1
                indexColor = 3
                statusNote = "Issue"
            End If
        End If
        statusPB = statusPB * 1

        If statusPB > 0 Then
            If pbDate = "" Or pbDate = 0 Or vendorId = "" Or totalAmount = 0 Then
                MsgBox "Please Select require field, try again."
                Exit Sub
            Else
                Dim access As String
                access = 0
                If totalAmount > 0 Then
                    '***Start save date To sales pb data***
                    Dim gl_id,gld_id As Integer

                    Set item_row         = activeWorkbook.Sheets("AddPurchaseBill").Range("I8:AK87")

                    Set pb_data_r        = ws_purchase_bill.Range("C8:AP8")
                    pb_id                = ws_purchase_bill.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    Set range_pb         = ws_pb_detail.Range("C8:AE8")
                    pb_detail_id         = ws_pb_detail.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    Set ws_gl_data       = ws_gl.Range("C8:Al8")
                    gl_id                = ws_gl.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    Set ws_gld_data      = ws_gld.Range("C8:AC8")
                    gld_id               = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    'insert sale pb
                    If (totalVat<=0) Then 
                        vatCal          = ""
                        vatSettingId    = ""
                        vatPercent      = 0
                        coaPurchaseVat   = ""
                    End If

                    pb_data_r.Cells(pb_id,1)  = pb_id 'id
                    pb_data_r.Cells(pb_id,2)  = pbCode 'pb no.
                    pb_data_r.Cells(pb_id,3)  = pbDate 'pb date
                    If (pbNetDays >=0 ) Then 
                        pb_data_r.Cells(pb_id,4)  = DateAdd("d", pbNetDays, pbDate) 'due date
                    Else
                        pb_data_r.Cells(pb_id,4)  =  ""
                    End If
                    pb_data_r.Cells(pb_id,5)  = purchaseRequestNo 'PR No. 
                    pb_data_r.Cells(pb_id,6)  = refinvoiceCode 'Ref# Invoice
                    If (pbRefInvoiceDate <> "" And pbRefInvoiceDate > 0) Then 
                        pb_data_r.Cells(pb_id,7)  = pbRefInvoiceDate'Ref. Invoice Date
                    Else
                        pb_data_r.Cells(pb_id,7)  = "" 'Ref. Invoice Date
                    End If
                    pb_data_r.Cells(pb_id,8)  = "" 'Purchase Receive Result Id
                    pb_data_r.Cells(pb_id,9)  = 1 'company Id
                    pb_data_r.Cells(pb_id,10) = branchId 'branch id
                    pb_data_r.Cells(pb_id,11) = locationGroupId 'location group id
                    pb_data_r.Cells(pb_id,12) = locationId 'location id
                    pb_data_r.Cells(pb_id,13) = paymentTermId 'payment term id
                    pb_data_r.Cells(pb_id,14) = vendorId 'vendor id
                    pb_data_r.Cells(pb_id,15) = vendorName 'vendor name
                    pb_data_r.Cells(pb_id,16) = coaPayable 'chart account ap id
                    pb_data_r.Cells(pb_id,17) = currencyId 'currency id
                    pb_data_r.Cells(pb_id,18) = exchangeRate 'exhange rate
                    pb_data_r.Cells(pb_id,19) = coaPurchaseVat 'Vat Chart Account Id
                    pb_data_r.Cells(pb_id,20) = vatPercent 'Vat Percent
                    pb_data_r.Cells(pb_id,21) = vatSettingId 'Vat Setting Id
                    pb_data_r.Cells(pb_id,22) = vatCal 'Vat Cal
                    pb_data_r.Cells(pb_id,23) = discountAmount 'Dis Amount
                    pb_data_r.Cells(pb_id,24) = discountPercent 'Dis​​ Percent
                    pb_data_r.Cells(pb_id,25) = totalAmount 'Sub Total
                    pb_data_r.Cells(pb_id,26) = totalVat 'total vat
                    pb_data_r.Cells(pb_id,27) = totalDiscount 'total discount
                    pb_data_r.Cells(pb_id,28) = totalDeposit 'total deposit
                    pb_data_r.Cells(pb_id,29) = totalBalance 'balance
                    pb_data_r.Cells(pb_id,30) = "" 'Principal
                    pb_data_r.Cells(pb_id,31) = "" 'Ship To
                    pb_data_r.Cells(pb_id,32) = "" 'Ship Telephone
                    pb_data_r.Cells(pb_id,33) = pbNote 'pb note
                    pb_data_r.Cells(pb_id,34) = createdDate 'created date
                    pb_data_r.Cells(pb_id,35) = 0 'Is Depsit Reference
                    pb_data_r.Cells(pb_id,36) = 0 'Is Update Cost
                    pb_data_r.Cells(pb_id,37) = statusPB 'pb status
                    pb_data_r.Cells(pb_id,38) = pbWeek 'week
                    pb_data_r.Cells(pb_id,39) = pbMonth 'month
                    pb_data_r.Cells(pb_id,40) = pbYear 'year
                    '***End save date To purchase data***

                    'insert general ledger
                    ws_gl_data.Cells(gl_id,1)  = gl_id 'id
                    ws_gl_data.Cells(gl_id,9)  = pb_id 'pb_id
                    ws_gl_data.Cells(gl_id,23) = pbDate 'date
                    ws_gl_data.Cells(gl_id,24) = pbCode 'reference
                    ws_gl_data.Cells(gl_id,25) = 0 'total_deposit
                    ws_gl_data.Cells(gl_id,27) = createdDate 'created
                    ws_gl_data.Cells(gl_id,28) = 1 'created_by
                    ws_gl_data.Cells(gl_id,31) = 1 'is_approve
                    ws_gl_data.Cells(gl_id,32) = 0 'is_depreciated
                    ws_gl_data.Cells(gl_id,33) = 0 'is_retained_earnings
                    ws_gl_data.Cells(gl_id,34) = 0 'deposit_type
                    ws_gl_data.Cells(gl_id,35) = 1 'is_active
                    ws_gl_data.Cells(gl_id,36) = pbWeek 'week
                    ws_gl_data.Cells(gl_id,37) = pbMonth 'month
                    ws_gl_data.Cells(gl_id,38) = pbYear 'year

                    'insert general ledger detail Accounts Receivable
                    ws_gld_data.Cells(gld_id,1)  = gld_id 'id
                    ws_gld_data.Cells(gld_id,2)  = gl_id 'general_ledger_id
                    ws_gld_data.Cells(gld_id,3)  = coaPayable 'chart_account_id
                    ws_gld_data.Cells(gld_id,4)  = 1 'company_id
                    ws_gld_data.Cells(gld_id,5)  = branchId 'branch id
                    ws_gld_data.Cells(gld_id,6)  = locationGroupId 'location group id
                    ws_gld_data.Cells(gld_id,7)  = locationId 'location id
                    ws_gld_data.Cells(gld_id,16) = "Bill" 'type
                    ws_gld_data.Cells(gld_id,17) = 0 'debit
                    ws_gld_data.Cells(gld_id,18) = totalBalance 'credit
                    ws_gld_data.Cells(gld_id,19) = "ICS: PB # " & pbCode 'memo
                    ws_gld_data.Cells(gld_id,20) = vendorId 'vendor_id
                    ws_gld_data.Cells(gld_id,24) = 1 'class_id
                    ws_gld_data.Cells(gld_id,25) = 1 'is_active
                    ws_gld_data.Cells(gld_id,26) = pbWeek 'week
                    ws_gld_data.Cells(gld_id,27) = pbMonth 'month
                    ws_gld_data.Cells(gld_id,28) = pbYear 'year

                    '*******Start save pb discount*******
                    If totalDiscount > 0 Then
                        'insert general ledger detail
                        gld_total_dis_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                        ws_gld_data.Cells(gld_total_dis_id,1)  = gld_total_dis_id 'id
                        ws_gld_data.Cells(gld_total_dis_id,2)  = gl_id 'general_ledger_id
                        ws_gld_data.Cells(gld_total_dis_id,3)  = coaPurchaseDiscount 'chart_account_id
                        ws_gld_data.Cells(gld_total_dis_id,4)  = 1 'company_id
                        ws_gld_data.Cells(gld_total_dis_id,5)  = branchId 'branch id
                        ws_gld_data.Cells(gld_total_dis_id,6)  = locationGroupId 'location group id
                        ws_gld_data.Cells(gld_total_dis_id,7)  = locationId 'location id
                        ws_gld_data.Cells(gld_total_dis_id,16) = "Bill" 'type
                        ws_gld_data.Cells(gld_total_dis_id,17) = 0 'debit
                        ws_gld_data.Cells(gld_total_dis_id,18) = totalDiscount 'credit
                        ws_gld_data.Cells(gld_total_dis_id,19) = "ICS: PB # " & pbCode & " Total Discount" 'memo
                        ws_gld_data.Cells(gld_total_dis_id,20) = vendorId 'vendor_id
                        ws_gld_data.Cells(gld_total_dis_id,24) = 1 'class_id
                        ws_gld_data.Cells(gld_total_dis_id,25) = 1 'is_active
                        ws_gld_data.Cells(gld_total_dis_id,26) = pbWeek 'week
                        ws_gld_data.Cells(gld_total_dis_id,27) = pbMonth 'month
                        ws_gld_data.Cells(gld_total_dis_id,28) = pbYear 'year
                        gld_total_dis_id = gld_total_dis_id + 1
                    End If
                    '*******End save pb discount*******

                    '*******Start save pb vat*******
                    If totalVat > 0 Then
                        'insert general ledger detail
                        gld_total_vat_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                        ws_gld_data.Cells(gld_total_vat_id,1)  = gld_total_vat_id 'id
                        ws_gld_data.Cells(gld_total_vat_id,2)  = gl_id 'general_ledger_id
                        ws_gld_data.Cells(gld_total_vat_id,3)  = coaPurchaseVat 'chart_account_id = 18 (VAT Output)
                        ws_gld_data.Cells(gld_total_vat_id,4)  = 1 'company_id
                        ws_gld_data.Cells(gld_total_vat_id,5)  = branchId 'branch id
                        ws_gld_data.Cells(gld_total_vat_id,6)  = locationGroupId 'location group id
                        ws_gld_data.Cells(gld_total_vat_id,7)  = locationId 'location id
                        ws_gld_data.Cells(gld_total_vat_id,16) = "Bill" 'type
                        ws_gld_data.Cells(gld_total_vat_id,17) = totalVat 'debit
                        ws_gld_data.Cells(gld_total_vat_id,18) = 0 'credit
                        ws_gld_data.Cells(gld_total_vat_id,19) = "ICS: PB # " & pbCode & " Total VAT" 'memo
                        ws_gld_data.Cells(gld_total_vat_id,20) = vendorId 'vendor_id
                        ws_gld_data.Cells(gld_total_vat_id,24) = 1 'class_id
                        ws_gld_data.Cells(gld_total_vat_id,25) = 1 'is_active
                        ws_gld_data.Cells(gld_total_vat_id,26) = pbWeek 'week
                        ws_gld_data.Cells(gld_total_vat_id,27) = pbMonth 'month
                        ws_gld_data.Cells(gld_total_vat_id,28) = pbYear 'year
                        gld_total_vat_id = gld_total_vat_id + 1
                    End If
                    ' *******End save pb vat*******

                    For a = 1 To item_row.Rows.Count
                        If item_row.Cells(a,1) <> "" Then 'old condition => WorksheetFunction.Count(rng.Rows(a)) <> 0
                            'insert sales pb detail
                            conversion           = 1
                            productId            = item_row.Cells(a,1)
                            productCode          = item_row.Cells(a,3)
                            productName          = Application.VLookup(productId * 1, rngProductList, 6, False)
                            qty                  = item_row.Cells(a, 6) * 1
                            qtyFree              = item_row.Cells(a, 7) * 1
                            productUoM           = MID(item_row.Cells(a,8),5,100)
                            productUomId         = MID(item_row.Cells(a,8),1,3) * 1
                            lotsNumber           = item_row.Cells(a,9)
                            expiredDate          = item_row.Cells(a,10)

                            newUnitCost          = item_row.Cells(a,12)
                            unitCostByUoM        = item_row.Cells(a,13)
                            disItem              = item_row.Cells(a,15)
                            totalCost            = item_row.Cells(a,17)
                            remark               = item_row.Cells(a,18)
                            If (item_row.Cells(a,19)>0) Then 
                                conversion      = item_row.Cells(a,19)
                            End If
                            smallValUoM          = item_row.Cells(a,20)
                            totalDisItem         = item_row.Cells(a,21)
                            defaultCost          = item_row.Cells(a,22)
                            itemType             = item_row.Cells(a,23) '1 : product, 2:service

                            checkExistProductInvTotal          = item_row.Cells(a,24)
                            checkExistProductInvTotalDetail    = item_row.Cells(a,25)

                            ' debug.Print "checkExistProductInvTotal=" & checkExistProductInvTotal & " ,checkExistProductInvTotalDetail=" & checkExistProductInvTotal
                            checkExistItemInvGroupTotal        = item_row.Cells(a,26)
                            checkExistItemInvGroupTotalDetail  = item_row.Cells(a,27)

                            stockByDateInvGroupTotalDetail     = item_row.Cells(a,28)
                            stockAvailableInvGroupTotal        = item_row.Cells(a,29)

                            'Updaet product cost
                            Set productRange     = ws_product.Range("B8:V8")
                            last_product_id      = ws_product.Cells(Rows.Count, "B").End(xlUp).Row

                            Set proCostHisRange  = ws_pro_cost_his.Range("C8:L8")
                            last_pro_cost_his    = ws_pro_cost_his.Cells(Rows.Count,"C").End(xlUp).Row - 6

                            If (newUnitCost > 0 And unitCostByUoM  <> newUnitCost) Then
                                proCostHisRange.Cells(last_pro_cost_his,1)  = last_pro_cost_his ' ID
                                proCostHisRange.Cells(last_pro_cost_his,2)  = pb_id ' Purchase Bill Id
                                proCostHisRange.Cells(last_pro_cost_his,3)  = pbCode ' Purchase Bill Code
                                proCostHisRange.Cells(last_pro_cost_his,4)  = productId  ' Product Id
                                proCostHisRange.Cells(last_pro_cost_his,5)  = productCode  ' Product Code
                                proCostHisRange.Cells(last_pro_cost_his,6)  = productName  ' Product Name
                                proCostHisRange.Cells(last_pro_cost_his,7)  = Application.VLookup(productId * 1, rngProductList, 8, False)  ' Product Uom
                                proCostHisRange.Cells(last_pro_cost_his,8)  = smallValUoM  ' Small Value Uom
                                proCostHisRange.Cells(last_pro_cost_his,9)  = unitCostByUoM * conversion  ' Old Cost
                                proCostHisRange.Cells(last_pro_cost_his,10) = newUnitCost * conversion   ' New Cost
                                proCostHisRange.Cells(last_pro_cost_his,11) = "Purchase"  ' Type
                                proCostHisRange.Cells(last_pro_cost_his,12) = createdDate ' Created
                                proCostHisRange.Cells(last_pro_cost_his,13) = 1' Created By
                                proCostHisRange.Cells(last_pro_cost_his,14) = 1 ' Status

                                For j = 8 To last_product_id
                                    If (ws_product.Cells(j,2).Value * 1 = productId * 1) Then
                                        'update unit cost product
                                        ws_product.Cells(j,13) = newUnitCost * conversion 'total qty sale 
                                        Exit For
                                    End If
                                Next j
                            End If

                            If (newUnitCost>0) Then 
                                unitCost = newUnitCost
                            Else
                                unitCost = unitCostByUoM
                            End If

                            totalAmountItem  = Format(totalCost + totalDisItem,"0.00")
                            qtyOrder         = (qty + qtyFree) / conversion '(qty + qtyFree) / smallValUoM / conversion
                            qtyOrderSmall    = (qty + qtyFree) * (smallValUoM / conversion)
                            costOrder        = Format((totalCost / (qty + qtyFree)) , "0.00") 
                            unitCost         = Format(unitCost,"0.00")
                            newUnitCost      = Format(newUnitCost,"0.00")
                            unitCostByUoM    = Format(unitCostByUoM,"0.00")
                            totalCost        = Format(totalCost,"0.00")   
                            ' lotsNumber,expiredDate
                            range_pb.Cells(pb_detail_id, 1)   = pb_detail_id 'id
                            range_pb.Cells(pb_detail_id, 2)   = pb_id 'pb id
                            range_pb.Cells(pb_detail_id, 3)   = "" 'Purchase Receive Id
                            range_pb.Cells(pb_detail_id, 4)   = pbCode 'pb number
                            range_pb.Cells(pb_detail_id, 5)   = pbDate 'pb date
                            range_pb.Cells(pb_detail_id, 6)   = vendorId 'vendor id
                            range_pb.Cells(pb_detail_id, 7)   = productId 'Product Id
                            range_pb.Cells(pb_detail_id, 8)   = productCode 'barcode product
                            range_pb.Cells(pb_detail_id, 9)   = productName 'item description/service
                            range_pb.Cells(pb_detail_id, 10)  = qty 'qty
                            range_pb.Cells(pb_detail_id, 11)  = qtyFree 'qty free
                            range_pb.Cells(pb_detail_id, 12)  = productUoM 'uom
                            range_pb.Cells(pb_detail_id, 13)  = productUomId 'uom id
                            range_pb.Cells(pb_detail_id, 14)  = conversion 'uom smallValUoM
                            range_pb.Cells(pb_detail_id, 15)  = "" 'Dis ID
                            range_pb.Cells(pb_detail_id, 16)  = disItem 'Dis Amount
                            range_pb.Cells(pb_detail_id, 17)  = "" 'Dis Percent
                            range_pb.Cells(pb_detail_id, 18)  = unitCost 'Default cost
                            range_pb.Cells(pb_detail_id, 19)  = costOrder 'New Unit Cost
                            range_pb.Cells(pb_detail_id, 20)  = unitCost 'unit cost
                            range_pb.Cells(pb_detail_id, 21)  = totalDisItem 'Total Dis Item
                            range_pb.Cells(pb_detail_id, 22)  = totalCost 'total cost
                            range_pb.Cells(pb_detail_id, 23)  = lotsNumber 'Lote Number
                            If (expiredDate > 0) Then 
                            range_pb.Cells(pb_detail_id, 24)  = expiredDate 'Expired date
                            Else
                            range_pb.Cells(pb_detail_id, 24)  = "" 'Expired date
                            End If
                            range_pb.Cells(pb_detail_id, 25)  = remark 'Note
                            range_pb.Cells(pb_detail_id, 26)  = 1 'status
                            range_pb.Cells(pb_detail_id, 27)  = pbWeek 'week
                            range_pb.Cells(pb_detail_id, 28)  = pbMonth 'month
                            range_pb.Cells(pb_detail_id, 29)  = pbYear 'year

                            Set range_inventory_valuation    = ws_pb_valuaction.Range("C8:AI8")
                            inv_valuation_id                 = ws_pb_valuaction.Cells(Rows.Count, "C").End(xlUp).Row - 6
                            ' insert inventory valucation
                            range_inventory_valuation.Cells(inv_valuation_id, 1)    = inv_valuation_id 'id
                            range_inventory_valuation.Cells(inv_valuation_id, 2)    = 1 'company_id
                            range_inventory_valuation.Cells(inv_valuation_id, 3)    = branchId 'branch_id
                            range_inventory_valuation.Cells(inv_valuation_id, 7)    = pb_id 'purchase bill id
                            range_inventory_valuation.Cells(inv_valuation_id, 8)    = pb_detail_id 'purchase bill detail id
                            range_inventory_valuation.Cells(inv_valuation_id, 13)   = "Bill" 'type=>Bill,Inventory Adjust,Invoice
                            range_inventory_valuation.Cells(inv_valuation_id, 14)   = pbCode 'reference
                            range_inventory_valuation.Cells(inv_valuation_id, 15)   = "" 'customer_id
                            range_inventory_valuation.Cells(inv_valuation_id, 16)   = vendorId 'vendor_id
                            range_inventory_valuation.Cells(inv_valuation_id, 17)   = pbDate 'date
                            range_inventory_valuation.Cells(inv_valuation_id, 18)   = productId 'product_id
                            range_inventory_valuation.Cells(inv_valuation_id, 19)   = qtyOrderSmall 'small_qty
                            range_inventory_valuation.Cells(inv_valuation_id, 20)   = qtyOrder 'qty
                            range_inventory_valuation.Cells(inv_valuation_id, 21)   = costOrder 'cost
                            range_inventory_valuation.Cells(inv_valuation_id, 22)   = 0 'price
                            range_inventory_valuation.Cells(inv_valuation_id, 23)   = 0 'on_hand
                            range_inventory_valuation.Cells(inv_valuation_id, 24)   = 0 'on_hand_small
                            range_inventory_valuation.Cells(inv_valuation_id, 25)   = costOrder 'avg_cost
                            range_inventory_valuation.Cells(inv_valuation_id, 26)   = 0 'asset_value
                            range_inventory_valuation.Cells(inv_valuation_id, 27)   = createdDate 'created
                            range_inventory_valuation.Cells(inv_valuation_id, 28)   = "" 'date_edited
                            range_inventory_valuation.Cells(inv_valuation_id, 29)   = 0 'is_refer_gm_id
                            range_inventory_valuation.Cells(inv_valuation_id, 30)   = 0 'avg_refer
                            range_inventory_valuation.Cells(inv_valuation_id, 31)   = 1 'is_var_cost
                            range_inventory_valuation.Cells(inv_valuation_id, 32)   = 0 'is_adjust_value
                            'range_inventory_valuation.Cells(inv_valuation_id, 33)   = 1 'is_active

                            'General Ledger Detail (Product Inventory)
                            gld_inventory_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                            ws_gld_data.Cells(gld_inventory_id,1)  = gld_inventory_id 'id
                            ws_gld_data.Cells(gld_inventory_id,2)  = gl_id 'general_ledger_id
                            ws_gld_data.Cells(gld_inventory_id,3)  = coaInventory 'chart_account_id =3
                            ws_gld_data.Cells(gld_inventory_id,4)  = 1 'company_id
                            ws_gld_data.Cells(gld_inventory_id,5)  = branchId 'branch id
                            ws_gld_data.Cells(gld_inventory_id,6)  = locationGroupId 'location group id
                            ws_gld_data.Cells(gld_inventory_id,7)  = locationId 'location id
                            ws_gld_data.Cells(gld_inventory_id,8)  = productId 'product_id
                            ws_gld_data.Cells(gld_inventory_id,11) = inv_valuation_id 'inventory_valuation_id
                            ws_gld_data.Cells(gld_inventory_id,12) = 0 'inventory_valuation_is_debit
                            ws_gld_data.Cells(gld_inventory_id,13) = pb_detail_id 'purchase bill detail id
                            ws_gld_data.Cells(gld_inventory_id,16) = "Bill" 'type=> Bill,Inventory Adjust,Invoice,Invoice Payment,Pay Bill,POS,Purchase Bill Payment,Receive Payment,Retained Earning 
                            ws_gld_data.Cells(gld_inventory_id,17) = totalAmountItem 'debit total amount before discount
                            ws_gld_data.Cells(gld_inventory_id,18) = 0 'credit
                            ws_gld_data.Cells(gld_inventory_id,19) = "ICS: PB # " & pbCode & " Product # " & productCode & " " &  productName 'memo
                            ws_gld_data.Cells(gld_inventory_id,20) = vendorId 'vendor_id
                            ws_gld_data.Cells(gld_inventory_id,24) = 1 'class_id
                            ws_gld_data.Cells(gld_inventory_id,25) = 1 'is_active
                            ws_gld_data.Cells(gld_inventory_id,26) = pbWeek 'week
                            ws_gld_data.Cells(gld_inventory_id,27) = pbMonth 'month
                            ws_gld_data.Cells(gld_inventory_id,28) = pbYear 'year
                            gld_inventory_id = gld_inventory_id + 1

                            If (totalDisItem > 0) Then
                                'insert general ledger detail Accounts Receivable
                                gld_discount_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                ws_gld_data.Cells(gld_discount_id,1)  = gld_discount_id 'id
                                ws_gld_data.Cells(gld_discount_id,2)  = gl_id 'general_ledger_id
                                ws_gld_data.Cells(gld_discount_id,3)  = coaPurchaseDiscount 'chart_account_id
                                ws_gld_data.Cells(gld_discount_id,4)  = 1 'company_id
                                ws_gld_data.Cells(gld_discount_id,5)  = branchId 'branch id
                                ws_gld_data.Cells(gld_discount_id,6)  = locationGroupId 'location group id
                                ws_gld_data.Cells(gld_discount_id,7)  = locationId 'location id
                                ws_gld_data.Cells(gld_discount_id,13) = pb_detail_id 'purchase bill detail id
                                ws_gld_data.Cells(gld_discount_id,16) = "Bill" 'type
                                ws_gld_data.Cells(gld_discount_id,17) = 0 'debit
                                ws_gld_data.Cells(gld_discount_id,18) = totalDisItem 'credit
                                ws_gld_data.Cells(gld_discount_id,19) = "ICS: PB # " & pbCode & " Product # " & productCode & " " & productName & " Discount" 'memo
                                ws_gld_data.Cells(gld_discount_id,20) = vendorId 'vendor_id
                                ws_gld_data.Cells(gld_discount_id,24) = 1 'class_id
                                ws_gld_data.Cells(gld_discount_id,25) = 1 'is_active
                                ws_gld_data.Cells(gld_discount_id,26) = pbWeek 'week
                                ws_gld_data.Cells(gld_discount_id,27) = pbMonth 'month
                                ws_gld_data.Cells(gld_discount_id,28) = pbYear 'year
                                gld_discount_id = gld_discount_id + 1
                            End If

                            inv_valuation_id = inv_valuation_id + 1
                            pb_detail_id     = pb_detail_id + 1

                            'insert inventory
                            Set inventory        = ws_pb.Range("C8:AA8")
                            inventory_id         = ws_pb.Cells(Rows.Count, "C").End(xlUp).Row - 6

                            inventory.Cells(inventory_id, 1)   = inventory_id 'id
                            inventory.Cells(inventory_id, 2)   = productId 'product id
                            inventory.Cells(inventory_id, 3)   = productCode 'product code
                            inventory.Cells(inventory_id, 4)   = productName 'product name
                            inventory.Cells(inventory_id, 5)   = pbDate 'inventory date
                            inventory.Cells(inventory_id, 6)   = locationId 'location id
                            inventory.Cells(inventory_id, 7)   = locationGroupId 'location group id
                            inventory.Cells(inventory_id, 8)   = "Purchase" 'inventory type (Inv Adj,Purchase,Sale,Void Sale...)
                            inventory.Cells(inventory_id, 10)   = vendorId 'vendor id
                            inventory.Cells(inventory_id, 15)  = pb_id 'Purchase Bill Id
                            inventory.Cells(inventory_id, 19)  = qtyOrderSmall  'Qty
                            inventory.Cells(inventory_id, 20)  = unitCost 'Unit Cost 
                            inventory.Cells(inventory_id, 21)  = 0 'Unit Price
                            inventory.Cells(inventory_id, 22)  = 1 'status
                            inventory.Cells(inventory_id, 23)  = lotsNumber 'Lots Number, 
                            If (expiredDate > 0) Then 
                            inventory.Cells(inventory_id, 24)  = expiredDate 'Expired Date
                            Else 
                            inventory.Cells(inventory_id, 24)  = "" 'Expired Date
                            End If
                            inventory.Cells(inventory_id, 25)  = createdDate 'created date
                            inventory.Cells(inventory_id, 26)  = pbWeek 'week
                            inventory.Cells(inventory_id, 27)  = pbMonth 'month
                            inventory.Cells(inventory_id, 28)  = pbYear 'year
                            inventory_id = inventory_id + 1

                            'Insert/Update 1_group_totals               
                            Set invGroupTotal        = ws_group_total.Range("C8:S8")
                            invGroupTotalId          = ws_group_total.Cells(Rows.Count, "C").End(xlUp).Row - 6
                            LastRowInvGroupTotal = ws_group_total.Cells(Rows.Count, "C").End(xlUp).Row
                            ' debug.Print "checkExistItemInvGroupTotal=" & checkExistItemInvGroupTotal & ",LastRowInvGroupTotal=" & LastRowInvGroupTotal
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
                                    'update inventory total
                                    If ws_group_total.Cells(i,3).Value * 1 = productId * 1 And ws_group_total.Cells(i,6).Value = locationGroupId And ws_group_total.Cells(i,7).Value = locationId And ws_group_total.Cells(i,8).Value = lotsNumber  And ws_group_total.Cells(i,9).Value = expiredDate Then
                                        qtyInvAdj    = ws_group_total.Cells(i, 11) * 1
                                        qtySale      = ws_group_total.Cells(i, 12) * 1
                                        qtyPos       = ws_group_total.Cells(i, 13) * 1
                                        qtyPb        = ws_group_total.Cells(i, 14) * 1
                                        qtyPr        = ws_group_total.Cells(i, 15) * 1
                                        qtySr        = ws_group_total.Cells(i, 16) * 1
                                        qtyToIn      = ws_group_total.Cells(i, 17) * 1
                                        qtyToOut     = ws_group_total.Cells(i, 18) * 1

                                        stockIn      = qtyOrderSmall + (qtyInvAdj + qtyPb + qtySr + qtyToIn)
                                        stockOut     = qtySale + qtyPos + qtyPr + qtyToOut '(qty + qtyFree) = New Qty Sales
                                        totalQty     = stockIn - stockOut

                                        ws_group_total.Cells(i,10)  = totalQty 'total qty
                                        ws_group_total.Cells(i,14)  = qtyPb + qtyOrderSmall 'total qty purchase bill
                                        Exit For
                                    End If
                                Next i
                            Else
                                'insert New record inventory group total invGroupTotal  ' lotsNumber,expiredDate
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
                                invGroupTotal.Cells(invGroupTotalId,8)   = qtyOrderSmall 'total qty ending
                                invGroupTotal.Cells(invGroupTotalId,9)   = 0 'total qty adjustment
                                invGroupTotal.Cells(invGroupTotalId,10)  = 0 'total qty sale
                                invGroupTotal.Cells(invGroupTotalId,11)  = 0 'Total Pos
                                invGroupTotal.Cells(invGroupTotalId,12)  = qtyOrderSmall 'Total PB
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
                            lastInvTotalGroupDetail  = ws_group_total_detail.Cells(Rows.Count, "C").End(xlUp).Row
                            ' debug.Print "checkExistItemInvGroupTotalDetail=" & checkExistItemInvGroupTotalDetail & ",lastInvTotalGroupDetail=" & lastInvTotalGroupDetail
                            If checkExistItemInvGroupTotalDetail > 0  Then
                                For j = 8 To lastInvTotalGroupDetail
                                    If (ws_group_total_detail.Cells(j,3).Value * 1 = productId * 1 And ws_group_total_detail.Cells(j,6).Value * 1 = locationGroupId And ws_group_total_detail.Cells(j,7).Value * 1 = locationId And ws_group_total_detail.Cells(j,8).Value = pbDate) Then
                                        'update  inventory total detail
                                        ws_group_total_detail.Cells(j,12) = ws_group_total_detail.Cells(j,12) * 1 + qtyOrderSmall 'total qty purchase bill 
                                        Exit For
                                    End If
                                Next j
                            Else
                                'insert New record inventory total detail
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,1)   = productId 'product id
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,2)   = productCode 'product code
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,3)   = productName 'product name
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,4)   = locationGroupId 'location id
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,5)   = locationId 'location id
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,6)   = pbDate 'date
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,7)   = 0 'total adjustment
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,8)   = 0 'total sale invoice
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,9)   = 0 'Total Pos
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,10)  = qtyOrderSmall 'Total Purchase Bill
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,11)  = 0 'Total Purchase Return
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,12)  = 0 'Total Sales Return
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,13)  = 0 'Total Transfer Order In
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,14)  = 0 'Total Transfer Order Out
                                invTotalGroupDetail.Cells(invTotalGroupDetailId,15)  = 0 'Total Order
                                invTotalGroupDetailId = invTotalGroupDetailId + 1
                            End If

                            'update & insert inventory total
                            Set range_inventory_total  = ws_pb_total.Range("C8:T8")
                            inventory_total_id         = ws_pb_total.Cells(Rows.Count, "C").End(xlUp).Row - 6
                            LastRowInvTotal            = ws_pb_total.Cells(Rows.Count, "C").End(xlUp).Row
                            ' debug.print "checkExistProductInvTotal=" & checkExistProductInvTotal
                            If checkExistProductInvTotal > 0 Then
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
                                    'update inventory total lotsNumber expiredDate
                                    ' debug.Print "product=>" & ws_pb_total.Cells(i,3).Value * 1 & "=" & productId * 1  & "=>lotsNumber : " &  ws_pb_total.Cells(i,6).Value  & "=" &  lotsNumber & "=>expiredDate : " &  ws_pb_total.Cells(i,7).Value  & "=" &  expiredDate 
                                    If ws_pb_total.Cells(i,3).Value * 1 = productId * 1 And ws_pb_total.Cells(i,6).Value = lotsNumber And ws_pb_total.Cells(i,7).Value = expiredDate  Then
                                        qtyInvAdj    = ws_pb_total.Cells(i, 9) * 1
                                        qtySale      = ws_pb_total.Cells(i, 10) * 1
                                        qtySaleFree  = ws_pb_total.Cells(i, 11) * 1
                                        qtyPos       = ws_pb_total.Cells(i, 12) * 1
                                        qtyPosFree   = ws_pb_total.Cells(i, 13) * 1
                                        qtyPb        = ws_pb_total.Cells(i, 14) * 1
                                        qtyPr        = ws_pb_total.Cells(i, 15) * 1
                                        qtySr        = ws_pb_total.Cells(i, 16) * 1
                                        qtySrFree    = ws_pb_total.Cells(i, 17) * 1
                                        qtyToIn      = ws_pb_total.Cells(i, 18) * 1
                                        qtyToOut     = ws_pb_total.Cells(i, 19) * 1

                                        stockIn      = qtyOrderSmall + qtyInvAdj + qtyPb + qtySr + qtySrFree + qtyToIn 
                                        stockOut     = qtySale + qtySaleFree + qtyPos + qtyPosFree + qtyPr + qtyToOut '(qty + qtyFree) = New Qty Sales
                                        totalQty     = stockIn - stockOut
                                        'range_inventory_total, ws_pb_total
                                        ws_pb_total.Cells(i,8)  = totalQty 'total qty
                                        ws_pb_total.Cells(i,14) = qtyPb + qtyOrderSmall 'total qty sale
                                        Exit For
                                    End If
                                Next i
                            ElseIf checkExistProductInvTotal = 0 Then
                                'insert New record inventory total
                                ' debug.Print "checkExistProductInvTotal=" & checkExistProductInvTotal & ",inventory_total_id=" & inventory_total_id
                                range_inventory_total.Cells(inventory_total_id,1)   = productId 'product id 
                                range_inventory_total.Cells(inventory_total_id,2)   = productCode 'product code
                                range_inventory_total.Cells(inventory_total_id,3)   = productName 'product name
                                range_inventory_total.Cells(inventory_total_id,4)   = lotsNumber 'lots number
                                If (expiredDate > 0) Then 
                                range_inventory_total.Cells(inventory_total_id,5)   = expiredDate 'expired date
                                Else
                                range_inventory_total.Cells(inventory_total_id,5)   = "" 'expired date
                                End If
                                range_inventory_total.Cells(inventory_total_id,6)   = qtyOrderSmall 'total qty ending
                                range_inventory_total.Cells(inventory_total_id,7)   = 0 'total qty adjustment
                                range_inventory_total.Cells(inventory_total_id,8)   = 0 'total qty sale
                                range_inventory_total.Cells(inventory_total_id,9)   = 0 'total qty free sale
                                range_inventory_total.Cells(inventory_total_id,10)  = 0 'Total Pos
                                range_inventory_total.Cells(inventory_total_id,11)  = 0 'Total Pos Free
                                range_inventory_total.Cells(inventory_total_id,12)  = qtyOrderSmall 'Total PB
                                range_inventory_total.Cells(inventory_total_id,13)  = 0 'Total PR
                                range_inventory_total.Cells(inventory_total_id,14)  = 0 'Total SR
                                range_inventory_total.Cells(inventory_total_id,15)  = 0 'Total SR Free
                                range_inventory_total.Cells(inventory_total_id,16)  = 0 'Total TO In
                                range_inventory_total.Cells(inventory_total_id,17)  = 0 'Total TO Out
                                range_inventory_total.Cells(inventory_total_id,18)  = 0 'Total Order
                                inventory_total_id = inventory_total_id + 1
                            End If

                            'update & insert inventory total detail
                            Set range_inventory_total_detail = ws_pb_total_detail.Range("C8:R8")
                            inv_total_detail_id              = ws_pb_total_detail.Cells(Rows.Count, "C").End(xlUp).Row - 6
                            lastRowInvTotalDetail            = ws_pb_total_detail.Cells(Rows.Count, "C").End(xlUp).Row

                            If checkExistProductInvTotalDetail > 0  Then
                                For j = 8 To lastRowInvTotalDetail
                                    If (ws_pb_total_detail.Cells(j,3).Value * 1 = productId * 1  AND ws_pb_total_detail.Cells(j,6).Value = locationId AND ws_pb_total_detail.Cells(j,7).Value = lotsNumber AND ws_pb_total_detail.Cells(j,8).Value = expiredDate And ws_pb_total_detail.Cells(j,9).Value = pbDate) Then
                                        'update  inventory total detail
                                        ws_pb_total_detail.Cells(j,13) = ws_pb_total_detail.Cells(j,13) * 1 + qtyOrderSmall 'total qty sale 
                                        Exit For
                                    End If
                                Next j
                            ElseIf checkExistProductInvTotalDetail = 0 Then
                                'insert New record inventory total detail
                                ' debug.Print "checkExistProductInvTotalDetail=" & checkExistProductInvTotalDetail & ",inv_total_detail_id=" & inv_total_detail_id
                                range_inventory_total_detail.Cells(inv_total_detail_id,1)   = productId 'product id
                                range_inventory_total_detail.Cells(inv_total_detail_id,2)   = productCode 'product code
                                range_inventory_total_detail.Cells(inv_total_detail_id,3)   = productName 'product name
                                range_inventory_total_detail.Cells(inv_total_detail_id,4)   = locationId 'location id
                                range_inventory_total_detail.Cells(inv_total_detail_id,5)   = lotsNumber 'lots number
                                If (expiredDate > 0) Then 
                                range_inventory_total_detail.Cells(inv_total_detail_id,6)   = expiredDate 'expired date
                                Else
                                range_inventory_total_detail.Cells(inv_total_detail_id,6)   = "" 'expired date
                                End If
                                range_inventory_total_detail.Cells(inv_total_detail_id,5)   = "" 'lots number
                                range_inventory_total_detail.Cells(inv_total_detail_id,6)   = "" 'expired date
                                range_inventory_total_detail.Cells(inv_total_detail_id,7)   = pbDate 'date
                                range_inventory_total_detail.Cells(inv_total_detail_id,8)   = 0 'total adjustment
                                range_inventory_total_detail.Cells(inv_total_detail_id,9)   = 0 'total sales invoice
                                range_inventory_total_detail.Cells(inv_total_detail_id,10)  = 0 'Total Pos
                                range_inventory_total_detail.Cells(inv_total_detail_id,11)  = qtyOrderSmall 'Total Purchase Bill
                                range_inventory_total_detail.Cells(inv_total_detail_id,12)  = 0 'Total Purchase Return
                                range_inventory_total_detail.Cells(inv_total_detail_id,13)  = 0 'Total Sales Return
                                range_inventory_total_detail.Cells(inv_total_detail_id,14)  = 0 'Total Transfer Order In
                                range_inventory_total_detail.Cells(inv_total_detail_id,15)  = 0 'Total Transfer Order Out
                                range_inventory_total_detail.Cells(inv_total_detail_id,16)  = 0 'Total Order
                                inv_total_detail_id = inv_total_detail_id + 1
                            End If
                        End If
                    Next a
                    '*******End save date To sales by item data*******
                    ' access = 1
                End If

                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_column_product").Value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_qty").Value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_qty_free").Value =""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_column_uom").Value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_column_new_unit_cost").Value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_discount_by_item").Value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_note").Value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_lots_number").value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_column_expired_date").value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_column_remark").Value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_purchase_request_no").Value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_ref_invoice_date").value = ""
                activeWorkbook.Sheets("AddPurchaseBill").Range("pb_ref_invoice").value = ""

                ws_purchase_bill.Select 'wbDatabase.
                ws_purchase_bill.Range("C" & ws_purchase_bill.Cells(Rows.Count, "C").End(xlUp).Row & ":" & "AP" & ws_purchase_bill.Cells(Rows.Count, "C").End(xlUp).Row).Select 'Select New insert pb row
                ' Run "PRint_Click"
                wbDatabase.Save
                OnEnd
                MsgBox "Purchase Bill " & pbCode & " saved successful."
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
End Sub

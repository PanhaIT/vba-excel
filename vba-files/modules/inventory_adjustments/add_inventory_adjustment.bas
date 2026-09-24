Attribute VB_Name = "add_inventory_adjustment"

Sub checkDatabaseFileAdj()
    ' Dim checkExistDb As Boolean
    ' Dim ws_file_setting As Worksheet
    ' Dim MyFSO As New FileSystemObject
    ' Dim databasePath, databaseFile, glFile, inventoryFile As String
    
    ' Set ws_file_setting = Sheets("FilesSetting")
    ' databasePath = ws_file_setting.Range("database_path").Value

    ' databaseFile = databasePath & "inventory_adjustments.xlsm"
    ' inventoryFile = databasePath & "inventories.xlsm"
    ' glFile = databasePath & "general_ledgers.xlsm"

    ' '1=Inventory Ajustment
    ' checkExistDb = checkDatabaseFile(1)
    ' If (checkExistDb = True) Then
    '     Application.ScreenUpdating = False
    '     Application.DisplayAlerts  = False

    '     Dim wb, wbDatabase As Workbook
    '     Dim ws As Worksheet
    '     Set ws = Sheets("InventoryAdj") 'specify here which worksheet to use
    '     Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
    '     ws.Range("R4").Copy Destination:=wbDatabase.Sheets("InvAdjData").Range("N6") 'pass text from text file to sheet excel
    '     wbDatabase.Close SaveChanges:=TRUE
    '     ' Application.CutCopyMode = False
    '     Application.ScreenUpdating = TRUE
    '     Application.DisplayAlerts  = TRUE
    ' Else
    
    ' End If
End Sub

Public Sub SaveAdjustStock()
    OnStart
    'Define variable
    Dim lotsNumber,refInvoice,checkRefStock,locationId,locationGroupId,invAdjMemo,createdBy,branchId,stockType,uom As String
    Dim a, b, i, j, LastRowInvTotal, lastRowInvTotalDetail As Long
    Dim invGroupTotal, invTotalGroupDetail, proCostHisRange, ws_gl_data, ws_gld_data, invValuation, rangeAddStock, rngInventoryAdj, InvAdjData, rngInventoryAdjDetail, rngInventory, rangInvTotalDetail As Range
    Dim LastRowInvGroupTotal, lastInvTotalGroupDetail, smallUomVal,conversion,invValId,gl_id,gld_id,gld_adj_detail_id,gld_detail_id As Integer
    Dim chartAccount,adjWeek,adjMonth,adjYear,invAdjId,invAdjDetailId,inventoryId,invTotalDetailId, statusAdjust, totalAdjQty As Integer
    Dim last_product_id, invGroupTotalId, invTotalGroupDetailId, productId, productCode, productName,qtyEnding, currentStock, qtyAdjustSmall, qtyInvTotal As Integer
    Dim newUnitCost, unitCostBigUom, bigQtyAdj,qtyOnHand,qtyOnHandSmall,costByUom, costSmallUom,totalCost As Double
    Dim qtySaleFree, qtyPosFree, qtySrFree, qtyInvAdj, qtySale, qtyPos, qtyPb, qtyPr, qtySr, qtyToIn, qtyToOut, stockIn, stockOut, totalQty As Integer
    Dim rngProductList,productRange As Range
    Dim adjustDate, createdDate, expiredDate As Date
    Dim ws_group_total, ws_group_total_detail, ws_product, ws_pro_cost_his, ws_gld, ws_gl, ws_adj,ws_adj_d,ws_inv,ws_inv_total,ws_inv_total_detail AS Worksheet
    Dim messageAlertBox AS VbMsgBoxResult
    Dim checkExistProductInvTotal,checkExistProductInvTotalDetail, checkExistItemInvGroupTotal,checkExistItemInvGroupTotalDetail,stockByDateInvGroupTotalDetail,stockAvailableInvGroupTotal As Integer
    Dim invAdjCode, remark As String

    invAdjCode      = ActiveSheet.Range("adj_code").Value
    adjustDate      = ActiveSheet.Range("adj_date").Value 'Format(ActiveSheet.Range("adj_date").Value, "yyyy/mm/dd")
    locationGroupId = ActiveSheet.Range("adj_location_group_id").Value 'adj_location
    stockType       = ActiveSheet.Range("adj_type").Value
    refInvoice      = ActiveSheet.Range("adj_ref_invoice").Value
    checkRefStock   = ActiveSheet.Range("adj_check_ref_invoice").value
    invAdjMemo      = ActiveSheet.Range("adj_note").Value
    
    If (ActiveSheet.Range("adj_branch").value <> "") Then 
        branchId    = MID(ActiveSheet.Range("adj_branch").value,1,2)*1
    Else
        MsgBox "Branch is requied, try again."
        Exit Sub
    End If

    createdBy       = ActiveSheet.Range("adj_stock_creator").Value
    totalAdjQty     = ActiveSheet.Range("adj_total_qty").Value
    chartAccount    = ActiveSheet.Range("adj_chart_account_id").Value
    adjWeek         = ActiveSheet.Range("adj_week").Value

    If (ActiveSheet.Range("adj_date").Value <> "") Then 
        adjMonth    = month(ActiveSheet.Range("adj_date").Value)
        adjYear     = year(ActiveSheet.Range("adj_date").Value)
    Else
        MsgBox "Inventory adjustment date is requied, try again."
        Exit Sub
    End If

    statusAdjust    = ActiveSheet.Range("adj_status").value
    createdDate     = Format(Now(), "yyyy/mm/dd hh:mm:ss")

    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws As Worksheet
    Dim checkExistDb As Boolean
    Dim ws_file_setting As Worksheet
    Dim MyFSO As New FileSystemObject
    Dim databasePath, databaseFile As String

    Set activeWorkbook = Workbooks("index.xlsm")

    messageAlertBox = MsgBox("Are you want to save this stock?", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Inventory Adjustment (Add New)")

    If (messageAlertBox = vbYes And checkDatabaseFile = TRUE) Then

        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
            databasePath    = ws_file_setting.Range("database_path").Value
            databaseFile    = databasePath & "database_kscm.xlsm"

        If (IsWorkbookOpen(CStr(databaseFile)) = FALSE) Then
            Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        End If

        Run resetAllFilter()
    
        If (checkRefStock = "exist") Then
            MsgBoxW ("This invoice stock already saved.")
            Exit Sub
        End If

        If (statusAdjust = 2) Then
            MsgBox "Inventory Adjustment already saved."
            Exit Sub
        ElseIf (statusAdjust = -1) Then
            MsgBox "Can't be save, Inventory Adjustment status is edit"
            Exit Sub
        ElseIf (adjustDate = "" or locationGroupId = ""  Or invAdjCode = "") Then 
            MsgBox "Please select require field, try again."
            Exit Sub
        End If

        totalCost = 0
        If statusAdjust = 1 Then
            Set ws_product            = activeWorkbook.Sheets("ProductList")
            Set rngProductList        = ws_product.Range("C8:AA" & ws_product.Cells(Rows.Count,"C").End(xlUp).Row)
            Set ws_pro_cost_his       = activeWorkbook.Sheets("ProductCostHistory")

            Set ws_adj                = wbDatabase.Sheets("InvAdjData")
            Set ws_adj_d              = wbDatabase.Sheets("InvAdjDetailData")
            Set ws_inv                = wbDatabase.Sheets("inventories")
            Set ws_invVal             = wbDatabase.Sheets("InventoryValuation")

            Set ws_gl                 = wbDatabase.sheets("GeneralLedger")
            Set ws_gld                = wbDatabase.Sheets("GeneralLedgerDetail")

            Set ws_group_total        = wbDatabase.Sheets(locationGroupId & "_group_totals")
            Set ws_group_total_detail = wbDatabase.Sheets(locationGroupId & "_group_total_details")

            Set rangeAddStock         = activeWorkbook.Sheets("InventoryAdj").Range("I8:AD87")

            Set rngInventoryAdj       = ws_adj.Range("C8:L8")
            invAdjId                  = ws_adj.Cells(Rows.Count, "C").End(xlUp).Row - 6

            Set rngInventoryAdjDetail = ws_adj_d.Range("C8:AB8")
            invAdjDetailId            = ws_adj_d.Cells(Rows.Count, "C").End(xlUp).Row - 6

            Set ws_gl_data            = ws_gl.Range("C8:Al8")
            gl_id                     = ws_gl.Cells(Rows.Count, "C").End(xlUp).Row - 6 'gl = general ledger

            Set ws_gld_data           = ws_gld.Range("C8:AC8")
            gld_id                    = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6 'gld = general ledger detail
      
            ' insert inventory adjustment
            rngInventoryAdj.Cells(invAdjId, 1)   = invAdjId 'id
            rngInventoryAdj.Cells(invAdjId, 2)   = branchId 'branch id
            rngInventoryAdj.Cells(invAdjId, 3)   = locationGroupId 'location group id
            rngInventoryAdj.Cells(invAdjId, 4)   = stockType 'stock type
            rngInventoryAdj.Cells(invAdjId, 5)   = chartAccount 'chart account id
            rngInventoryAdj.Cells(invAdjId, 6)   = adjustDate 'adjustment date
            rngInventoryAdj.Cells(invAdjId, 7)   = invAdjCode 'adjustment no.
            rngInventoryAdj.Cells(invAdjId, 8)   = refInvoice 'reference no.
            rngInventoryAdj.Cells(invAdjId, 9)   = invAdjMemo 'memo
            rngInventoryAdj.Cells(invAdjId, 10)  = 2 'status

            'insert general ledger
            ws_gl_data.Cells(gl_id,1)  = gl_id 'id
            ws_gl_data.Cells(gl_id,2)  = invAdjId 'inv_adj_id
            ws_gl_data.Cells(gl_id,23) = adjustDate 'date
            ws_gl_data.Cells(gl_id,24) = invAdjCode 'reference
            ws_gl_data.Cells(gl_id,25) = 0 'total_deposit
            ws_gl_data.Cells(gl_id,27) = createdDate 'created
            ws_gl_data.Cells(gl_id,28) = 1 'created_by
            ws_gl_data.Cells(gl_id,29) = 1 'is_sys
            ws_gl_data.Cells(gl_id,30) = 0 'is_adj
            ws_gl_data.Cells(gl_id,31) = 1 'is_approve
            ws_gl_data.Cells(gl_id,32) = 0 'is_depreciated
            ws_gl_data.Cells(gl_id,33) = 0 'is_retained_earnings
            ws_gl_data.Cells(gl_id,34) = 0 'deposit_type
            ws_gl_data.Cells(gl_id,35) = 1 'is_active
            ws_gl_data.Cells(gl_id,36) = adjWeek 'week
            ws_gl_data.Cells(gl_id,37) = adjMonth 'month
            ws_gl_data.Cells(gl_id,38) = adjYear 'year

            For b = 1 To rangeAddStock.Rows.Count
                If WorksheetFunction.Count(rangeAddStock.Rows(b)) <> 0 Then
                    productId         = rangeAddStock.Cells(b,1)
                    productCode       = rangeAddStock.Cells(b,3)
                    productName       = Application.VLookup(CInt(productId), rngProductList, 6, False) 'MID(Split(rangeAddStock.Cells(b,4),"#")(0),7,100)
                    uom               = rangeAddStock.Cells(b,5)
                    locationId        = MID(rangeAddStock.Cells(b,6),1,3) * 1
                    currentStock      = rangeAddStock.Cells(b,7)
                    qtyAdjustSmall    = rangeAddStock.Cells(b,8)
                    qtyEnding         = rangeAddStock.Cells(b,9)
                    newUnitCost       = rangeAddStock.Cells(b,10)
                    lotsNumber        = rangeAddStock.Cells(b,11)
                    expiredDate       = rangeAddStock.Cells(b,12)
                    remark            = rangeAddStock.Cells(b,13)
                    conversion        = rangeAddStock.Cells(b,14)
                    smallUomVal       = rangeAddStock.Cells(b,15)
                    costByUom         = rangeAddStock.Cells(b,16)
                    unitCostBigUom    = rangeAddStock.Cells(b,17)
  
                    checkExistProductInvTotal          = rangeAddStock.Cells(b,18) 'Check exist product item in 1_inventory_totals group by product_id, lots_number,expired Date
                    checkExistProductInvTotalDetail    = rangeAddStock.Cells(b,19) 'Check exist product item in 1_inventory_total_details group by product_id, location_id, lots_number,expired Date and date

                    checkExistItemInvGroupTotal        = rangeAddStock.Cells(b,20) 'Check exist product item in 1_group_totals group by product_id, lots_number,expired Date,location_group_id,location_id
                    checkExistItemInvGroupTotalDetail  = rangeAddStock.Cells(b,21) 'Check exist product item in 1_inventory_total_details group by product_id,location_group_id location_id and date

                    stockByDateInvGroupTotalDetail     = rangeAddStock.Cells(b,22) 'total stock from period start to inventery adjustment date of 1_group_total_details
                    stockAvailableInvGroupTotal        = qtyAdjustSmall 'total stock available in warehouse of 1_group_totals

                    'Updaet product cost
                    Set productRange     = ws_product.Range("B8:V8")
                    Set proCostHisRange  = ws_pro_cost_his.Range("C8:L8")
                    last_product_id      = ws_product.Cells(Rows.Count, "B").End(xlUp).Row
                    last_pro_cost_his    = ws_pro_cost_his.Cells(Rows.Count,"C").End(xlUp).Row - 6

                    If (newUnitCost > 0 And costByUom  <> newUnitCost) Then
                        proCostHisRange.Cells(last_pro_cost_his,1)  = last_pro_cost_his ' ID
                        proCostHisRange.Cells(last_pro_cost_his,2)  = invAdjId ' Purchase Bill Id
                        proCostHisRange.Cells(last_pro_cost_his,3)  = invAdjCode ' Purchase Bill Code
                        proCostHisRange.Cells(last_pro_cost_his,4)  = productId  ' Product Id
                        proCostHisRange.Cells(last_pro_cost_his,5)  = productCode  ' Product Code
                        proCostHisRange.Cells(last_pro_cost_his,6)  = productName  ' Product Name
                        proCostHisRange.Cells(last_pro_cost_his,7)  = Application.VLookup(CInt(productId), rngProductList, 8, False)  ' Product Uom
                        proCostHisRange.Cells(last_pro_cost_his,8)  = smallUomVal  ' Small Value Uom
                        proCostHisRange.Cells(last_pro_cost_his,9)  = costByUom * conversion  ' Old Cost
                        proCostHisRange.Cells(last_pro_cost_his,10) = newUnitCost * conversion   ' New Cost
                        proCostHisRange.Cells(last_pro_cost_his,11) = "Inventory Adjust"  ' Type
                        proCostHisRange.Cells(last_pro_cost_his,12) = createdDate ' Created
                        proCostHisRange.Cells(last_pro_cost_his,13) = 1' Created By
                        proCostHisRange.Cells(last_pro_cost_his,14) = 1 ' Status

                        For j = 8 To last_product_id
                            If (CInt(ws_product.Cells(j,3).Value) = CInt(productId)) Then
                                'update unit cost product
                                ws_product.Cells(j,16) = newUnitCost * conversion 'total qty sale 
                                Exit For
                            End If
                        Next j
                    End If

                    qtyOnHandSmall  = 0
                    qtyOnHand       = 0
                    If (smallUomVal > 0) Then
                        bigQtyAdj  = Format(qtyAdjustSmall / smallUomVal , "0.000000000") ' bigQtyAdj  = bigQtyAdj /smallUomVal/ conversion
                    End If

                    If (newUnitCost > 0) Then
                        unitCost = newUnitCost
                        totalCost = Format(bigQtyAdj * newUnitCost,"0.00")
                    Else
                        unitCost = costByUom
                        totalCost = Format(bigQtyAdj * unitCostBigUom,"0.00")
                    End If
                    
                    unitCost    = Format(unitCost,"0.00")
                    newUnitCost = Format(newUnitCost,"0.00")
                    costByUom   = Format(costByUom,"0.00")

                    If (qtyAdjustSmall < 0 ) Then 
                        totalCost = totalCost * -1
                    End If

                    'insert inventory adjustment detail
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 1)   = invAdjDetailId 'id
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 2)   = invAdjId 'inv adj id
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 3)   = refInvoice 'ref invoice stock id
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 4)   = adjustDate 'rngInventory date
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 5)   = locationId 'location id
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 6)   = lotsNumber 'lots number
                    If (expiredDate > 0) Then 
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 7)   = expiredDate 'expired date
                    Else
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 7)   = "" 'expired date
                    End If
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 8)   = stockType 'created by
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 9)   = productId 'item id
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 10)  = productCode 'item code
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 11)  = productName 'product name

                    rngInventoryAdjDetail.Cells(invAdjDetailId, 12)  = qtyEnding 'ending qty
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 13)  = currentStock 'count qty
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 14)  = qtyAdjustSmall 'adjust qty

                    rngInventoryAdjDetail.Cells(invAdjDetailId, 15)  = uom 'uom
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 16)  = conversion 'conversion
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 17)  = newUnitCost 'new unit cost
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 18)  = unitCost 'average cost
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 19)  = totalCost 'asset value
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 20)  = createdDate 'created date
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 21)  = createdBy 'created by
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 22)  = remark
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 23)  = 1 'status
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 24)  = adjWeek 'week
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 25)  = adjMonth 'month
                    rngInventoryAdjDetail.Cells(invAdjDetailId, 26)  = adjYear 'year
                    invAdjDetailId = invAdjDetailId + 1

                    'insert inventory valucation
                    Set invValuation  = ws_invVal.Range("C8:AI8")
                    invValId          = ws_invVal.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    invValuation.Cells(invValId, 1)    = invValId 'id
                    invValuation.Cells(invValId, 2)    = 1 'company_id
                    invValuation.Cells(invValId, 3)    = branchId 'branch_id
                    invValuation.Cells(invValId, 4)    = invAdjId 'adjustment id
                    invValuation.Cells(invValId, 13)   = "Inventory Adjust" 'type=> Bill,Inventory Adjust,Invoice
                    invValuation.Cells(invValId, 14)   = refInvoice 'reference
                    invValuation.Cells(invValId, 17)   = adjustDate 'date
                    invValuation.Cells(invValId, 18)   = productId 'product_id
                    invValuation.Cells(invValId, 19)   = qtyAdjustSmall 'small_qty
                    invValuation.Cells(invValId, 20)   = bigQtyAdj 'qty
                    invValuation.Cells(invValId, 21)   = unitCost 'cost
                    invValuation.Cells(invValId, 22)   = 0 'price
                    invValuation.Cells(invValId, 23)   = qtyOnHand 'qtyOnHand 'on_hand
                    invValuation.Cells(invValId, 24)   = qtyOnHandSmall 'qtyOnHandSmall 'on_hand_small
                    invValuation.Cells(invValId, 25)   = unitCost 'avg_cost
                    invValuation.Cells(invValId, 26)   = qtyOnHand * unitCost 'asset_value
                    invValuation.Cells(invValId, 27)   = createdDate 'created
                    invValuation.Cells(invValId, 29)   = 0 'is_refer_gm_id
                    invValuation.Cells(invValId, 30)   = 0 'avg_refer
                    invValuation.Cells(invValId, 31)   = 1 'is_var_cost
                    invValuation.Cells(invValId, 32)   = 0 'is_adjust_value
                    invValuation.Cells(invValId, 33)   = 1 'is_active

                    'General Ledger Detail (Inventory Adjustment Detial)
                    gld_adj_detail_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                    ws_gld_data.Cells(gld_adj_detail_id,1)  = gld_adj_detail_id 'id
                    ws_gld_data.Cells(gld_adj_detail_id,2)  = gl_id 'general_ledger_id
                    ws_gld_data.Cells(gld_adj_detail_id,3)  = 3 'chart_account_id => inventoy asset account chartAccount
                    ws_gld_data.Cells(gld_adj_detail_id,4)  = 1 'company_id
                    ws_gld_data.Cells(gld_adj_detail_id,5)  = branchId 'branch id
                    ws_gld_data.Cells(gld_adj_detail_id,6)  = locationGroupId 'location group id
                    ws_gld_data.Cells(gld_adj_detail_id,7)  = locationId 'location id
                    ws_gld_data.Cells(gld_adj_detail_id,8)  = productId 'product_id
                    ws_gld_data.Cells(gld_adj_detail_id,11) = invValId 'inventory_valuation_id
                    ws_gld_data.Cells(gld_adj_detail_id,16) = "Inventory Adjust" 'type=> Bill,Inventory Adjust,Invoice,Invoice Payment,Pay Bill,POS,Purchase Bill Payment,Receive Payment,Retained Earning
                    If (qtyAdjustSmall>0) Then
                        ws_gld_data.Cells(gld_adj_detail_id,12) = 1 'inventory_valuation_is_debit
                        ws_gld_data.Cells(gld_adj_detail_id,17) = totalCost 'debit
                        ws_gld_data.Cells(gld_adj_detail_id,18) = 0 'credit
                    Else
                        ws_gld_data.Cells(gld_adj_detail_id,12) = 0 'inventory_valuation_is_debit
                        ws_gld_data.Cells(gld_adj_detail_id,17) = 0 'debit
                        ws_gld_data.Cells(gld_adj_detail_id,18) = totalCost  'credit
                    End If

                    ws_gld_data.Cells(gld_adj_detail_id,19) = "ICS: Inventory adjustment for product # " & productCode & " " & productName 'memo
                    ws_gld_data.Cells(gld_adj_detail_id,20) = customerId 'customer_id
                    ws_gld_data.Cells(gld_adj_detail_id,24) = 1 'class_id
                    ws_gld_data.Cells(gld_adj_detail_id,25) = 1 'is_active
                    ws_gld_data.Cells(gld_adj_detail_id,26) = adjWeek 'week
                    ws_gld_data.Cells(gld_adj_detail_id,27) = adjMonth 'month
                    ws_gld_data.Cells(gld_adj_detail_id,28) = adjYear 'year
                    gld_adj_detail_id = gld_adj_detail_id + 1

                    'General Ledger Detail (Inventory Adjustment Detial)
                    gld_detail_id = ws_gld.Cells(Rows.Count, "C").End(xlUp).Row - 6
                    ws_gld_data.Cells(gld_detail_id,1)  = gld_adj_detail_id 'id
                    ws_gld_data.Cells(gld_detail_id,2)  = gl_id 'general_ledger_id
                    ws_gld_data.Cells(gld_detail_id,3)  = chartAccount 'chart_account_id => Equity
                    ws_gld_data.Cells(gld_detail_id,4)  = 1 'company_id
                    ws_gld_data.Cells(gld_detail_id,5)  = branchId 'branch id
                    ws_gld_data.Cells(gld_detail_id,6)  = locationGroupId 'location group id
                    ws_gld_data.Cells(gld_detail_id,7)  = locationId 'location id
                    ws_gld_data.Cells(gld_detail_id,8)  = productId 'product_id
                    ws_gld_data.Cells(gld_detail_id,11) = invValId 'inventory_valuation_id
                    ws_gld_data.Cells(gld_detail_id,16) = "Inventory Adjust" 'type

                    If (qtyAdjustSmall>0) Then
                        ws_gld_data.Cells(gld_detail_id,12) = 0 'inventory_valuation_is_debit
                        ws_gld_data.Cells(gld_detail_id,17) = 0 'debit
                        ws_gld_data.Cells(gld_detail_id,18) = totalCost 'credit
                    Else
                        ws_gld_data.Cells(gld_detail_id,12) = 1 'inventory_valuation_is_debit
                        ws_gld_data.Cells(gld_detail_id,17) = totalCost 'debit
                        ws_gld_data.Cells(gld_detail_id,18) = 0 'credit
                    End If

                    ws_gld_data.Cells(gld_detail_id,19) = "ICS: Inventory adjustment for product # " & productCode & " " &  productName 'memo
                    ws_gld_data.Cells(gld_detail_id,20) = customerId 'customer_id
                    ws_gld_data.Cells(gld_detail_id,24) = 1 'class_id
                    ws_gld_data.Cells(gld_detail_id,25) = 1 'is_active
                    ws_gld_data.Cells(gld_detail_id,26) = adjWeek 'week
                    ws_gld_data.Cells(gld_detail_id,27) = adjMonth 'month
                    ws_gld_data.Cells(gld_detail_id,28) = adjYear 'year
                    gld_detail_id = gld_detail_id + 1

                    invValId = invValId + 1

                    'insert inventory 
                    Set rngInventory    = ws_inv.Range("C8:AD8")
                    inventoryId         = ws_inv.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    rngInventory.Cells(inventoryId, 1)   = inventoryId 'id
                    rngInventory.Cells(inventoryId, 2)   = productId 'product id
                    rngInventory.Cells(inventoryId, 3)   = productCode 'product code
                    rngInventory.Cells(inventoryId, 4)   = productName 'product name
                    rngInventory.Cells(inventoryId, 5)   = adjustDate 'rngInventory date
                    rngInventory.Cells(inventoryId, 6)   = locationId 'location id
                    rngInventory.Cells(inventoryId, 7)   = locationGroupId 'location group id
                    rngInventory.Cells(inventoryId, 8)   = "Inv Adj" 'rngInventory type (Inv Adj,Purchase,Sale,Void Sale...)
                    rngInventory.Cells(inventoryId, 11)  = invAdjId 'Adj Id
                    rngInventory.Cells(inventoryId, 19)  = qtyAdjustSmall 'Qty
                    rngInventory.Cells(inventoryId, 20)  = unitCost 'Unit Cost
                    rngInventory.Cells(inventoryId, 21)  = 0 'Unit Price
                    rngInventory.Cells(inventoryId, 22)  = 1 'status
                    rngInventory.Cells(inventoryId, 23)  = lotsNumber 'lots number
                    If (expiredDate > 0) Then 
                    rngInventory.Cells(inventoryId, 24)  = expiredDate 'expired date
                    Else
                    rngInventory.Cells(inventoryId, 24)  = "" 'expired date
                    End If
                    rngInventory.Cells(inventoryId, 25)  = createdDate 'created date
                    rngInventory.Cells(inventoryId, 26)  = adjWeek 'week
                    rngInventory.Cells(inventoryId, 27)  = adjMonth 'month
                    rngInventory.Cells(inventoryId, 28)  = adjYear 'year
                    inventoryId = inventoryId + 1

                    'Insert/Update 1_group_totals
                    Set invGroupTotal        = ws_group_total.Range("C8:S8")
                    invGroupTotalId          = ws_group_total.Cells(Rows.Count, "C").End(xlUp).Row - 6
                    
                    LastRowInvGroupTotal = ws_group_total.Cells(Rows.Count, "C").End(xlUp).Row
                    ' debug.print "checkExistItemInvGroupTotal=" & checkExistItemInvGroupTotal
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
                            ' update inventory total
                            If ws_group_total.Cells(i,3).Value * 1 = productId * 1 And ws_group_total.Cells(i,6).Value = locationGroupId And ws_group_total.Cells(i,7).Value = locationId And ws_group_total.Cells(i,8).Value = lotNumber  And ws_group_total.Cells(i,9).Value = expiredDate Then
                                ' debug.Print "***********************Update group total , productId=" & productId
                                qtyInvAdj    = ws_group_total.Cells(i, 11) * 1
                                qtySale      = ws_group_total.Cells(i, 12) * 1
                                qtyPos       = ws_group_total.Cells(i, 13) * 1
                                qtyPb        = ws_group_total.Cells(i, 14) * 1
                                qtyPr        = ws_group_total.Cells(i, 15) * 1
                                qtySr        = ws_group_total.Cells(i, 16) * 1
                                qtyToIn      = ws_group_total.Cells(i, 17) * 1
                                qtyToOut     = ws_group_total.Cells(i, 18) * 1

                                stockIn      = qtyAdjustSmall + (qtyInvAdj + qtyPb + qtySr + qtyToIn)
                                stockOut     = qtySale + qtyPos + qtyPr + qtyToOut '(qty + qtyFree) = New Qty Sales
                                totalQty     = stockIn - stockOut

                                ws_group_total.Cells(i,10)  = totalQty 'total qty
                                ws_group_total.Cells(i,11)  = qtyInvAdj + qtyAdjustSmall 'total qty adjustment
                                Exit For
                            End If
                        Next i
                    Else
                        ' debug.Print "***********************insert new group total , productId=" & productId
                        'insert New record inventory group total invGroupTotal  ' lotNumber,expiredDate
                        invGroupTotal.Cells(invGroupTotalId,1)   = productId 'product id 
                        invGroupTotal.Cells(invGroupTotalId,2)   = productCode 'product code
                        invGroupTotal.Cells(invGroupTotalId,3)   = productName 'product name
                        invGroupTotal.Cells(invGroupTotalId,4)   = locationGroupId 'location group id
                        invGroupTotal.Cells(invGroupTotalId,5)   = locationId 'location id
                        If (lotNumber<> "") Then 
                        invGroupTotal.Cells(invGroupTotalId,6)   = lotNumber 'lote number
                        Else
                        invGroupTotal.Cells(invGroupTotalId,6)   = "" 'lote number
                        End If
                        If (expiredDate >0) Then 
                        invGroupTotal.Cells(invGroupTotalId,7)   = expiredDate 'expired date
                        Else
                        invGroupTotal.Cells(invGroupTotalId,7)   = "" 'expired date
                        End If
                        invGroupTotal.Cells(invGroupTotalId,8)   = qtyAdjustSmall 'total qty ending
                        invGroupTotal.Cells(invGroupTotalId,9)   = qtyAdjustSmall 'total qty adjustment
                        invGroupTotal.Cells(invGroupTotalId,10)  = 0 'total qty sale
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

                    lastInvTotalGroupDetail = ws_group_total_detail.Cells(Rows.Count, "C").End(xlUp).Row
                    If checkExistItemInvGroupTotalDetail > 0  Then
                        For j = 8 To lastInvTotalGroupDetail
                            If (ws_group_total_detail.Cells(j,3).Value * 1 = productId * 1 And ws_group_total_detail.Cells(j,6).Value * 1 = locationGroupId And ws_group_total_detail.Cells(j,7).Value * 1 = locationId And ws_group_total_detail.Cells(j,8).Value = adjustDate) Then
                                'update  inventory total detail
                                ws_group_total_detail.Cells(j,9) = ws_group_total_detail.Cells(j,9) * 1 + qtyAdjustSmall 'total qty adjustment
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
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,6)   = adjustDate 'date
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,7)   = qtyAdjustSmall 'total adjustment
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,8)   = 0 'total sale invoice
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,9)   = 0 'Total Pos
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,10)  = 0 'Total Purchase Bill
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,11)  = 0 'Total Purchase Return
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,12)  = 0 'Total Sales Return
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,13)  = 0 'Total Transfer Order In
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,14)  = 0 'Total Transfer Order Out
                        invTotalGroupDetail.Cells(invTotalGroupDetailId,15)  = 0 'Total Order
                        invTotalGroupDetailId = invTotalGroupDetailId + 1
                    End If

                    'update and insert inventory total
                    Set ws_inv_total          = wbDatabase.Sheets(locationId & "_inventory_totals")
                    Set ws_inv_total_detail   = wbDatabase.Sheets(locationId & "_inventory_total_details")
                    
                    Set invTotal  = ws_inv_total.Range("C8:T8")
                    invTotalId    = ws_inv_total.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    LastRowInvTotal = ws_inv_total.Cells(Rows.Count, "C").End(xlUp).Row
                    If (checkExistProductInvTotal > 0 ) Then
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
                            If ws_inv_total.Cells(i,3).Value * 1 = productId * 1 And ws_inv_total.Cells(i,6).Value = lotsNumber And ws_inv_total.Cells(i,7).Value = expiredDate Then
                                'update inventory total
                                qtyInvAdj    = ws_inv_total.Cells(i, 9) * 1
                                qtySale      = ws_inv_total.Cells(i, 10) * 1
                                qtySaleFree  = ws_inv_total.Cells(i, 11) * 1
                                qtyPos       = ws_inv_total.Cells(i, 12) * 1
                                qtyPosFree   = ws_inv_total.Cells(i, 13) * 1
                                qtyPb        = ws_inv_total.Cells(i, 14) * 1
                                qtyPr        = ws_inv_total.Cells(i, 15) * 1
                                qtySr        = ws_inv_total.Cells(i, 16) * 1
                                qtySrFree    = ws_inv_total.Cells(i, 17) * 1
                                qtyToIn      = ws_inv_total.Cells(i, 16) * 1
                                qtyToOut     = ws_inv_total.Cells(i, 17) * 1

                                stockIn      = qtyAdjustSmall + (qtyInvAdj + qtyPb + qtySr + qtySrFree + qtyToIn)
                                stockOut     = qtySale + qtySaleFree + qtyPos + qtyPosFree + qtyPr + qtyToOut '(qty + qtyFree) = New Qty Sales
                                totalQty     = stockIn - stockOut

                                ws_inv_total.Cells(i,8) = totalQty 'total qty
                                ws_inv_total.Cells(i,9) = ws_inv_total.Cells(i,9) + qtyAdjustSmall 'total qty stock
                            End If
                        Next i
                    Else
                        'insert new record inventory total
                        invTotal.Cells(invTotalId,1)   = productId 'product id 
                        invTotal.Cells(invTotalId,2)   = productCode 'product code
                        invTotal.Cells(invTotalId,3)   = productName 'product name
                        invTotal.Cells(invTotalId,4)   = lotsNumber 'lots number
                        If (expiredDate > 0) Then 
                        invTotal.Cells(invTotalId,5)   = expiredDate 'expired date
                        Else
                        invTotal.Cells(invTotalId,5)   = "" 'expired date
                        End If
                        invTotal.Cells(invTotalId,6)   = qtyAdjustSmall 'total qty ending
                        invTotal.Cells(invTotalId,7)   = qtyAdjustSmall 'total qty adjustment
                        invTotal.Cells(invTotalId,8)   = 0 'total qty sale
                        invTotalId = invTotalId + 1
                    End If

                    'update and insert total detail
                    Set rangInvTotalDetail = ws_inv_total_detail.Range("C8:R8")
                    invTotalDetailId       = ws_inv_total_detail.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    lastRowInvTotalDetail = ws_inv_total_detail.Cells(Rows.Count, "C").End(xlUp).Row
                    If checkExistProductInvTotalDetail > 0  Then
                        For j = 8 To lastRowInvTotalDetail
                            If (ws_inv_total_detail.Cells(j,3).Value * 1 = productId * 1 AND ws_inv_total_detail.Cells(j,6).Value = locationId AND ws_inv_total_detail.Cells(j,7).Value = lotsNumber AND ws_inv_total_detail.Cells(j,8).Value = expiredDate AND ws_inv_total_detail.Cells(j,9).Value = adjustDate) Then
                                ' update inventory total detail
                                ws_inv_total_detail.Cells(j,10) = ws_inv_total_detail.Cells(j,10) + qtyAdjustSmall 'total qty 
                            End If
                        Next j
                    Else
                        'insert new record inventory total detail
                        rangInvTotalDetail.Cells(invTotalDetailId,1)   = productId 'product id
                        rangInvTotalDetail.Cells(invTotalDetailId,2)   = productCode 'product code
                        rangInvTotalDetail.Cells(invTotalDetailId,3)   = productName 'product name
                        rangInvTotalDetail.Cells(invTotalDetailId,4)   = locationId 'product name
                        rangInvTotalDetail.Cells(invTotalDetailId,5)   = lotsNumber 'lots number
                        If (expiredDate > 0) Then 
                        rangInvTotalDetail.Cells(invTotalDetailId,6)   = expiredDate 'expired date
                        Else
                        rangInvTotalDetail.Cells(invTotalDetailId,6)   = "" 'expired date
                        End If
                        rangInvTotalDetail.Cells(invTotalDetailId,7)   = adjustDate 'date
                        rangInvTotalDetail.Cells(invTotalDetailId,8)   = qtyAdjustSmall 'total adjustment
                        rangInvTotalDetail.Cells(invTotalDetailId,9)   = 0 'total sale invoice
                        invTotalDetailId = invTotalDetailId + 1
                    End If
                End If
            Next b

            activeWorkbook.Sheets("InventoryAdj").Range("adj_item").Value = ""
            activeWorkbook.Sheets("InventoryAdj").Range("adj_column_uom").Value = ""
            activeWorkbook.Sheets("InventoryAdj").Range("adj_ref_invoice").Value = ""
            activeWorkbook.Sheets("InventoryAdj").Range("adj_note").Value = ""
            activeWorkbook.Sheets("InventoryAdj").Range("adj_column_location").Value = ""
            activeWorkbook.Sheets("InventoryAdj").Range("adj_column_current_qty").Value = ""
            activeWorkbook.Sheets("InventoryAdj").Range("adj_column_lots_number").Value = ""
            activeWorkbook.Sheets("InventoryAdj").Range("adj_column_remark").Value = ""
            activeWorkbook.Sheets("InventoryAdj").Range("adj_column_expired_date").Value = ""
            
            ' clearInvAdj (obj) ' clear add stock form
            ws_adj_d.Select
            ' invAdjId = ws_adj_d.Cells(Rows.Count, "C").End(xlUp).Row
            ActiveSheet.Range("C" & ws_adj_d.Cells(Rows.Count, "C").End(xlUp).Row & ":" & "AB" & ws_adj_d.Cells(Rows.Count, "C").End(xlUp).Row).Select
            wbDatabase.Save
            OnEnd
            MsgBox "Inventory Adjustment successful saved."
            Exit Sub
        End If
    Else
        exit sub
    End If
End Sub

Public Function clearInvAdj(obj)
    ActiveSheet.Range("adj_column_input_stock").Value = ""
    ActiveSheet.Range("adj_item").Value = ""
    ' ActiveSheet.Range("adj_date").Value = ""
    ActiveSheet.Range("adj_ref_invoice").Value = ""
    ActiveSheet.Range("adj_note").Value = ""
End Function

Private Sub adjCurrentDate()
    ActiveSheet.Range("adj_date").Value = Format(Now(), "yyyy/mm/dd")
End Sub
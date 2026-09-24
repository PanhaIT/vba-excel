VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' 1.Id
' 2.Inv Adj Id
' 3.Ref.Inv.Stock
' 4.Inv Date 
' 5.Location
' 6.Lots Number
' 7.Expired date
' 8.Stock Type 
' 9.Item Id
' 10.Item code
' 11.Item Description
' 12.Ending Qty
' 13.Count Qty
' 14.Adjust Qty (+/-)
' 15.UoM
' 16.Conversion
' 17.New Unit Cost
' 18.Avg Cost
' 19.Asset Value
' 20.Created 
' 21.Created By
' 22.Remark
' 23.Status
' 24.Week
' 25.Month
' 26.Year

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

Private Sub invAdjComboBox_Change()
    ' OnStart
    'loop variable
    Dim i, j, k, lastRow As Integer

    'workbook/worksheet variable
    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_gld, ws_gl, ws, ws_inv_data, ws_inv_item_data,ws_file_setting, ws_print_adj, ws_warehouse, ws_product As Worksheet

    'branch variable
    Dim branchId, branchName,branchNameKh As String
    Dim branchListRange As Range
    Dim bankName, accountName, accountNo, branchAddressKh, branchTel, branchEmail As String

    'product variable
    Dim productId, qty, qtyFree, productCode, productName, productUoM As String
    Dim rngUom, item_row,rngProductList As Range
    Dim lastProductRow, smallValUom As Integer

    'inventory adjustment variable
    Dim adjDate As Date
    Dim statusNote, priceTypeId, discountTypeId, adjCode, adjNote,warehouseName,referenceInvAdj As String
    Dim AdjId, netDays, paymentTermId, conversion, qtyOrder, qtyOrderSmall, invValId, invId, invDetailId, exchangeRate, warehouseId, lastLocationGroupId As Integer
    Dim lastAdjId, countAdjDetailItem, vatChartAccId, lastAdjDetailId As Integer
    Dim rngInvAdjDataList, rowInvAdjustmentCode, rngAdjDetail, rangePrintInvItem, rangeInvDetailId  As Range
    Dim criteriaRange1, criteriaRange2, rngWarehouse As Range
    Dim countItem As Integer

    'setting variable
    Dim companyId As Integer
    Dim vatPercent, vatSettingId, vatCal, chartAccIncome,lastRowLocation As Integer
    Dim companyNameKh, companyNameEn,companyTinNumber, databasePath, databaseFile,locationName As String
    Dim companyListRange, paymentTermRange,rngLocation As Range

    'amount variable
    Dim totalCost, totalDeposit, totalDiscount, totalAmount, grandTotalAmount, unitCost, totalDisItem, disItem, totalVat As Double

    Set activeWorkbook     = Workbooks("index.xlsm")
    Set ws_print_adj       = activeWorkbook.Sheets("PrintInvAdjustment")
    Set rangePrintInvItem  = ws_print_adj.Range("print_adj_items")
    Set ws_warehouse       = activeWorkbook.Sheets("LocationGroup")
    Set ws_product         = activeWorkbook.Sheets("ProductList")

    If (Me.invAdjComboBox.Value <> "Please select adjustment code" And Me.invAdjComboBox.Value <> "") Then 'checkDatabaseFile = TRUE And
        rangePrintInvItem.ClearContents
        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath        = ws_file_setting.Range("database_path").Value
        databaseFile        = databasePath & "database_kscm.xlsm"

        ' If (IsWorkbookOpen(CStr(databaseFile)) = FALSE) Then
        '     ' Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        '     MsgBox "Please open database file befor select Adj code."
        ' End If

        Set wbDatabase       = Workbooks("database_kscm.xlsm")
        Set ws_inv_data      = wbDatabase.Sheets("InvAdjData")
        Set ws_inv_item_data = wbDatabase.Sheets("InvAdjDetailData")

        lastAdjId            = ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row
        lastAdjDetailId      = ws_inv_item_data.Cells(Rows.Count, "C").End(xlUp).Row
        lastLocationGroupId  = ws_warehouse.Cells(Rows.Count, "C").End(xlUp).Row
        lastProductRow       = ws_product.Cells(Rows.Count,"C").End(xlUp).Row

        Set rngInvAdjDataList    = ws_inv_data.Range("C8:L" & lastAdjId)
        Set rangeInvDetailId     = ws_inv_item_data.Range("C8:C" & lastAdjDetailId)
        Set rngAdjDetail         = ws_inv_item_data.Range("C8:AB" & lastAdjDetailId)
        Set rngWarehouse         = ws_warehouse.Range("C8:K" & lastLocationGroupId)
        Set rngProductList       = ws_product.Range("C8:Z" & lastProductRow)

        countAdjDetailItem  = Application.CountA(rangeInvDetailId)
        AdjId               = CInt(Split(Me.invAdjComboBox.Value, "#")(1))
        adjCode             = Split(Me.invAdjComboBox.Value, "#")(0)
       
        companyId       = 1
        adjDate         = Application.VLookup(AdjId, rngInvAdjDataList, 6, False) 'Format(, "mm/dd/yyyy")
        branchId        = Application.VLookup(AdjId, rngInvAdjDataList, 2, False)
        adjNote         = Application.VLookup(AdjId, rngInvAdjDataList, 9, False)
        warehouseId     = Application.VLookup(AdjId, rngInvAdjDataList, 3, False)
        referenceInvAdj = Application.VLookup(AdjId, rngInvAdjDataList, 8, False)

        lastRowLocation = activeWorkbook.Sheets("Location").Cells(Rows.Count,"C").End(xlUp).Row
        ' Set companyListRange  = activeWorkbook.Sheets("Setting").Range("company_list")
        Set branchListRange   = activeWorkbook.Sheets("Branch").Range("branch_list")
        Set rngUom            = activeWorkbook.Sheets("UoM").Range("uom_list")
        Set rngLocation       = activeWorkbook.Sheets("Location").Range("C8:N" & lastRowLocation)

        ' companyNameKh    = Application.VLookup(companyId, companyListRange, 2, False)
        ' companyNameEn    = Application.VLookup(companyId, companyListRange, 3, False)
        ' companyTinNumber = Application.VLookup(companyId, companyListRange, 5, False)
        branchName       = Application.VLookup(branchId, branchListRange, 5, False)
        branchNameKh     = Application.VLookup(branchId, branchListRange, 4, False)
        branchAddressKh  = Application.VLookup(branchId, branchListRange, 22, False)
        branchTel        = Application.VLookup(branchId, branchListRange, 18, False)
        branchEmail      = Application.VLookup(branchId, branchListRange, 20, False)
        warehouseName    = Application.VLookup(warehouseId,rngWarehouse,4,False)

        'adjustment Information
        ws_print_adj.Range("print_adj_code").Value = adjCode
        ws_print_adj.Range("print_adj_date").Value = adjDate
        ws_print_adj.Range("print_adj_note").Value = adjNote
        ' ws_print_adj.Range("print_adj_company_name_kh").Value = companyNameKh
        ' ws_print_adj.Range("print_adj_company_name_en").Value = companyNameEn
        ' ws_print_adj.Range("print_adj_company_tin").Value = companyTinNumber
        ws_print_adj.Range("print_adj_warehouse").Value = warehouseName 
        ws_print_adj.Range("print_adj_reference").Value = referenceInvAdj 
        
        'Branch information 
        ws_print_adj.Range("print_adj_branch_address_kh").Value = branchAddressKh
        ws_print_adj.Range("print_adj_branch_telephone").Value = branchTel
        ws_print_adj.Range("print_adj_branch_email").Value = branchEmail

        ws_print_adj.Range("C27" & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = False 'Show all row in range print invoice
        ' debug.print "countAdjDetailItem=" & countAdjDetailItem

        If AdjId > 0 And countAdjDetailItem > 0 Then
            k = 0
            grandTotalAmount = 0
            For j = 1 To countAdjDetailItem
                If (rngAdjDetail.Cells(j, 2) = AdjId) Then
                    ' disItem             = rngAdjDetail.Cells(j,6)
                    ' totalDisItem        = rngAdjDetail.Cells(j,6)
                    productId       = rngAdjDetail.Cells(j, 9)
                    productCode     = rngAdjDetail.Cells(j, 10)
                    productName     = rngAdjDetail.Cells(j, 11)           'MID(Split(rngAdjDetail.Cells(j, 7),"#")(0), 7, 100)
                    qty             = rngAdjDetail.Cells(j, 14)
                    ' qtyFree         = rngAdjDetail.Cells(j, 11)
                    productUoM      = Application.VLookup(Mid(rngAdjDetail.Cells(j, 15), 1, 3) * 1, rngUom, 5, False)
                    unitCost        = rngAdjDetail.Cells(j, 18)
                    totalCost       = rngAdjDetail.Cells(j, 19)
                    locationName    = Application.VLookup(rngAdjDetail.Cells(j, 5), rngLocation, 4, False)
                    smallValUom     = Application.Vlookup(productId,rngProductList,18,False)
                    conversion      = rngAdjDetail.Cells(j, 16)

                    Set criteriaRange1 = ws_inv_item_data.Range("D8:D" & countAdjDetailItem + 7) 'column invoice id
                    Set criteriaRange2 = ws_inv_item_data.Range("Y8:Y" & countAdjDetailItem + 7) 'column status
                    countItem          = Application.COUNTIFS(criteriaRange1,AdjId,criteriaRange2,1)
                    ' debug.Print "countItem=" & countItem

                    If (qty > 0) Then
                        k = k + 1
                        'set condition to manage page
                        If (countItem >= 26 AND countItem <= 104) Then '26, 52, 78, 104
                            If (k Mod 26 = 0) Then  '=> limit 26 rows per page
                                ' Debug.Print "***************k=" & k
                                k = k + 1
                            End If
                        Else
                            If (k Mod 13 = 0) Then ' k = 13 => limit 12 rows per page
                                ' Debug.Print "***************k=" & k
                                k = k + 14
                            End If
                        End If
                        ' debug.Print "k=" & k
                        rangePrintInvItem.Cells(k, 1) = k
                        rangePrintInvItem.Cells(k, 2) = productCode
                        rangePrintInvItem.Cells(k, 3) = productName
                        rangePrintInvItem.Cells(k, 6) = locationName
                        rangePrintInvItem.Cells(k, 7) = (qty/smallValUom) * conversion
                        rangePrintInvItem.Cells(k, 8) = productUoM
                        rangePrintInvItem.Cells(k, 9) = unitCost
                        rangePrintInvItem.Cells(k, 10) = totalCost
                        grandTotalAmount = grandTotalAmount + totalCost
                    End If
                End If
            Next j
            lastRow = k + 27
            ws_print_adj.Range("C" & lastRow & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
        End If
        ws_print_adj.Range("print_adj_total_cost_usd").Value = grandTotalAmount
    Else
        rangePrintInvItem.ClearContents
        ws_print_adj.Range("C27:C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
    End If
    
    ' OnEnd

End Sub

Private Sub Worksheet_Activate()
    ' OnStart

    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_print_adj, ws_inv_data, ws_file_setting As Worksheet
    Dim rowInvAdjustmentCode, rowAdjId As Range
    Dim i, CountaResultCell, lastAdjId As Integer
    Dim databasePath, databaseFile As String

    Set activeWorkbook = Workbooks("index.xlsm")
    Set ws_print_adj    = activeWorkbook.Sheets("PrintInvAdjustment")

    ws_print_adj.invAdjComboBox.Clear
    ws_print_adj.invAdjComboBox.AddItem "Please select adjustment code"

    If (checkDatabaseFile = True) Then
        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath        = ws_file_setting.Range("database_path").Value
        databaseFile        = databasePath & "database_kscm.xlsm"

        ' If (IsWorkbookOpen(databaseFile) = False) Then
        '     'Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        ' End If

        Set wbDatabase  = Workbooks("database_kscm.xlsm")
        Set ws_inv_data = wbDatabase.Sheets("InvAdjData")
        lastAdjId       = ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row - 7
    
        Set rowInvAdjustmentCode = ws_inv_data.Range("I8:I" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row)
        Set rowAdjId             = ws_inv_data.Range("C8:C" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row)
        For i = 1 To lastAdjId
            If rowInvAdjustmentCode.Cells(i) <> "" Then
                ws_print_adj.invAdjComboBox.AddItem rowInvAdjustmentCode(i) & " #" & rowAdjId(i)
            End If
        Next i
    Else
        ' MsgBox "Database file doesn't exist."
        ' Exit Sub
    End If
    
    ' OnEnd
End Sub

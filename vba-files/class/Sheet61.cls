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

' 1.Id
' 2.PB ID
' 3.Purchase Receive Id
' 4.PB Code
' 5.PB Date
' 6.Vendor Id
' 7.Product Id
' 8.Product Code
' 9.Product Name
' 10.Qty
' 11.Qty Free
' 12.UoM
' 13.Uom Id
' 14.Conversion
' 15.Discount Id
' 16.Discount Amount
' 17.Discount Percent
' 18.Default Cost
' 19.New Unit Cost
' 20.Unit Cost
' 21.Total Dis Item
' 22.Total Cost
' 23.Lote Number
' 24.Expired date
' 25.Remark
' 26.Status
' 27.Week
' 28.Month
' 29.Year

' 1.Id
' 2.PB ID
' 3.PB Code
' 4.PB Date
' 5.Vendor Id
' 6.Service Id
' 7.Service Code
' 8.Service Name
' 9.Qty
' 10.Qty Free
' 11.UoM
' 12.Uom Id
' 13.Conversion
' 14.Discount Id
' 15.Discount Amount
' 16.Discount Percent
' 17.Default Cost
' 18.New Cost
' 19.Total Dis Item
' 20.Unit Cost
' 21.Total Cost
' 22.Remark
' 23.Status
' 24.Week
' 25.Month
' 26.Year

Private Sub PurchaseBillComboBox_Change()
    ' OnStart
    'loop variable
    Dim i, j, k,lastRow As Integer

    'workbook/worksheet variable
    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_gld, ws_gl, ws, ws_inv_data, ws_inv_item_data,ws_file_setting, ws_print_pb,ws_pb_service As Worksheet

    'vendor variable
    Dim vendorId, vendorName As String
    Dim vendorListRange As Range
    Dim vendorNameEn, vendorNameKh, vendorAddressKh, vendorTel, vendorEmail, vendorTinNumber As String

    'branch variable
    Dim branchId, branchName As String
    Dim branchListRange As Range
    Dim bankName, accountName, accountNo, branchAddressKh, branchTel, branchEmail,socialNetwork As String

    'product variable
    Dim productId, qty, qtyFree, productCode, productName, productUoM As String
    Dim rngUom, item_row As Range

    'purchase invoice variable
    Dim pbDate, pbDueDate As Date
    Dim statusNote, priceTypeId, discountTypeId, PbCode, pbNote As String
    Dim purchaseId, netDays, paymentTermId, conversion, qtyOrder, qtyOrderSmall, invValId, pbNo, invId, invDetailId, exchangeRate As Integer
    Dim lastPurchaseId, countPurchaseDetailItem,countPurchaseServiceItem, vatChartAccId, lastPbDetailId As Integer
    Dim rngPbServiceId, rngPurchaseBillList, rowPbCode, rngPbDetail, rangePrintInvItem, rangeInvDetailId  As Range
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
        ws_print_pb.Range("U27:AW143").Value = ""
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
        Set ws_pb_service    = wbDatabase.Sheets("PurchaseBillService")
        ' lastPurchaseId       = ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row
        ' lastPbDetailId       = ws_inv_item_data.Cells(Rows.Count, "C").End(xlUp).Row

        Set rngPurchaseBillList  = ws_inv_data.Range("C8:AP" & ws_inv_data.Cells(Rows.Count, "C").End(xlUp).Row)
        Set rangeInvDetailId     = ws_inv_item_data.Range("C8:C" & ws_inv_item_data.Cells(Rows.Count, "C").End(xlUp).Row)
        Set rngPbServiceId       = ws_pb_service.Range("C8:C" & ws_pb_service.Cells(Rows.Count, "C").End(xlUp).Row)
        Set rngPbDetail          = ws_inv_item_data.Range("C8:Z" & ws_inv_item_data.Cells(Rows.Count, "C").End(xlUp).Row)
        
        ' countPurchaseDetailItem  = Application.CountA(rangeInvDetailId)
        ' countPurchaseServiceItem = Application.CountA(rngPbServiceId)
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
        socialNetwork    = Application.VLookup(branchId, branchListRange, 21, False)
        If IsError(companyNameKh) Then
            companyNameKh = ""
        End If

        'pb Information
        ws_print_pb.Range("print_pb_id").Value = purchaseId
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
        ws_print_pb.Range("print_pb_social_network").Value = socialNetwork

        'vendor Information
        ws_print_pb.Range("print_pb_vendor_name_kh").Value = vendorNameKh
        ws_print_pb.Range("print_pb_vendor_name_en").Value = vendorNameEn
        ws_print_pb.Range("print_pb_vendor_address_kh").Value = vendorAddressKh
        ws_print_pb.Range("print_pb_vendor_tel").Value = vendorTel
        ws_print_pb.Range("print_pb_vendor_email").Value = vendorEmail
        ws_print_pb.Range("print_pb_vendor_tin_number").Value = vendorTinNumber

        ws_print_pb.Range("C27" & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = False 'Show all row in range print invoice

        If purchaseId > 0 Then
            ' U19:AT20
            ' U21:AW22
            ' U27:AW27
            ' *******advanced filter item data************
            Dim rngPbItem As Range
            Dim rngSearchCriteria As Range
            Dim rngPbItemNoHead As Range

            'define the database and criteria ranges
            If ws_inv_item_data.FilterMode Then ws_inv_item_data.ShowAllData
            On Error Resume Next
            ActiveSheet.ShowAllData
            On Error GoTo 0

            Set rngPbItem = ws_inv_item_data.Range("C7:AE" & ws_inv_item_data.Cells(Rows.Count, "C").End(xlUp).Row) 'database range
            Set rngSearchCriteria = ws_print_pb.Range("U21:AW22") 
            
            'filter the database using the criteria
            rngPbItem.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=rngSearchCriteria, Unique:=False

            ' 2. Exclude the header row using Offset(1, 0) and Resize
            On Error Resume Next ' Handles cases where 0 rows match the filter
                Set rngPbItemNoHead = rngPbItem.Offset(1, 0).Resize(rngPbItem.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0

            If Not rngPbItemNoHead Is Nothing Then
                rngPbItemNoHead.Copy Destination:=ws_print_pb.Range("U27:AW27")
                countPurchaseDetailItem  = Application.CountA(ws_print_pb.Range("U27:U" & ws_print_pb.Cells(Rows.Count, "U").End(xlUp).Row))
            Else
                countPurchaseDetailItem = 0
            End If

            ' 3. Copy and paste to the next empty row on the destination sheet
            If ws_inv_item_data.FilterMode Then ws_inv_item_data.ShowAllData

            '*******advanced filter service data************

            Dim rngPbService As Range
            Dim rngSearchServiceCriteria As Range
            Dim rngPbServiceNoHead As Range
            Dim rowCount, colCount As Long

            Set rngPbService = ws_pb_service.Range("C7:AB" & ws_pb_service.Cells(Rows.Count, "C").End(xlUp).Row) 'database range
            Set rngSearchServiceCriteria = ws_print_pb.Range("U19:AT20")

            'filter the database using the criteria
            rngPbService.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=rngSearchServiceCriteria, Unique:=False

            ' 2. Exclude the header row using Offset(1, 0) and Resize
            On Error Resume Next ' Handles cases where 0 rows match the filter
                Set rngPbServiceNoHead = rngPbService.Offset(1, 0).Resize(rngPbService.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0

            '************Resize array service much sales invoice by item column************

            ' Exit if no visible rows remain after filtering
            If ws_pb_service.FilterMode Then ws_pb_service.ShowAllData 'clear filter on the source sheet

            If rngPbServiceNoHead Is Nothing Then
                rowCount = 0
            Else
                Dim sourceArr, desArr, arrHeader1, arrHeader2 As Variant
                Dim r,c, insertAfterCol1 ,numNewCols1 ,insertAfterCol2,numNewCols2 As Long

                ' Convert the multi-area visible range into a 2D Variant array
                sourceArr = VisibleRangeToArray(rngPbServiceNoHead)

                ' 1. Define the new headers to insert after Column U (Column 21)
                arrHeader1          = Array("Purchase Receive Id") 'insert header after column D (column 4)
                arrHeader2          = Array("Lote Number", "Expired Date") 'insert header after column T (column 20)

                insertAfterCol1     = 2 'insert between columns D and E (after column D)
                insertAfterCol2     = 22 'insert between columns T and U (after column T)
                numNewCols1         = UBound(arrHeader1) - LBound(arrHeader1) + 1
                numNewCols2         = UBound(arrHeader2) - LBound(arrHeader2) + 1

                ' 2. Load existing data into a 1-based 2D array (assumes headers are in Row 1)
                rowCount            = UBound(sourceArr, 1) 'total row service data
                colCount            = UBound(sourceArr, 2) 'total column service data
                
                ReDim desArr(1 To rowCount, 1 To colCount + numNewCols1 + numNewCols2)
                
                ' 4. Loop through rows and columns to map data and insert headers
                For r = 1 To rowCount
                    For c = 1 To colCount
                        If c > insertAfterCol1 And c <= insertAfterCol2 Then
                            desArr(r, c + numNewCols1) = sourceArr(r, c)
                            If c = 3 Then
                                desArr(r, insertAfterCol1 + 1) = "" 'Purchase Receive Id
                            End If
                        ElseIf c >= insertAfterCol2 Then
                            ' Shift columns after the second insertion point
                            desArr(r, c + numNewCols1 + numNewCols2) = sourceArr(r, c)
                            If c = 23 Then
                                desArr(r, insertAfterCol2 + 1) = "" 'lots number
                            End If
                            If c = 24 Then
                                desArr(r, insertAfterCol2 + 2) = "" 'expired date
                            End If
                        Else
                            desArr(r, c) = sourceArr(r, c)
                        End If
                    Next c
                Next r
                ws_print_pb.Range("U" & countPurchaseDetailItem + 27).Resize(rowCount, colCount + numNewCols1 + numNewCols2).Value = desArr
            End If

            If ws_pb_service.FilterMode Then ws_pb_service.ShowAllData 'clear filter on the source sheet
            lastRow = countPurchaseDetailItem + rowCount + 27
            
            ' ws_print_pb.Range("C" & lastRow & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
        End If
    Else
        ws_print_pb.Range("U27:AW143").Value = "" ' rangePrintInvItem.ClearContents
        ' ws_print_pb.Range("C27:C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
    End If

    ' OnEnd
End Sub

Private Sub Worksheet_Activate()
    ' OnStart
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

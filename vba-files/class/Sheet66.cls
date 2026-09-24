VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'*************Quote Data*************
' 1 . No.
' 2 . Quote Code
' 3 . Quote Date
' 4 . Company Id
' 5 . Branch Id
' 6 . Customer Id
' 7 . Customer Name
' 8 . Service Type
' 9 . Currency ID
' 10 . Price Type Id
' 11 . Exchange Rate
' 12 . Vat Percent
' 13 . Vat Setting ID
' 14 . Vat Cal
' 15 . Dis Amount
' 16 . Dis​​ Percent
' 17 . Sub Total
' 18 . Total VAT
' 19 . Total Discount
' 20 . Total Deposit
' 21 . Total Amount
' 22 . Ship To
' 23 . Ship Telephone
' 24 . Note
' 25 . Created
' 26 . Created By
' 27 . Approved
' 28 . Approved By
' 29 . Is Close
' 30 . Is Approve
' 31 . Status
' 32 . Week
' 33 . Month
' 34 . Year

'*********Quotation Detail************
' 1 . Id
' 2 . Quote ID
' 3 . Quote No.
' 4 . Quote Date
' 5 . Customer Id
' 6 . Product Id
' 7 . Product Code
' 8 . Product Name
' 9 . Qty
' 10 . Qty Free
' 11 . UoM
' 12 . Qty Uom Id
' 13 . Conversion
' 14 . Discount Id
' 15 . Discount Amount
' 16 . Discount Percent
' 17 . Unit Cost
' 18 . Unit Price
' 19 . Dis Total Price
' 20 . Total Price
' 21 . Status
' 22 . Week
' 23 . Month
' 24 . Year

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

Private Sub QuotationComboBox_Change()
    ' OnStart
    'workbook , worksheet variable
    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_product, ws_file_setting, ws_print_quote, ws_gld, ws_gl, ws, ws_quote_data, ws_quote_item_data As Worksheet

    'loop variable
    Dim i, j, k, m, LastRow As Integer

    'customer variable
    Dim lastRowCustomer As Integer
    Dim customerName, customerId, customerNameEn, customerNameKh, customerAddressKh, customerTel, customerEmail, customerTinNumber As String

    'Quote variable
    Dim quoteDate, QuoteDueDate As Date
    Dim statusNote, priceTypeId, discountTypeId, quoteCode, quoteNote As String
    Dim vatCal, vatPercent, quoteId, netDays, conversion, invId, invDetailId, exchangeRate, lastQuoteDetailId As Integer

    'amount variable
    Dim totalAmount, unitPrice, totalPrice, totalDeposit, totalDiscount, grandTotalAmount, totalDisItem, disItem, totalVat As Double

    'Product variable
    Dim lastRowFileSetting, countQuoteDetailItem, countQuoteNo, qtyStock, productId, qty, qtyFree, convertion, lastRowUom,lastProductRow As Integer
    Dim productCode, productName, productUoM,specProduct As String
    Dim rngProductList, rngUom, rangeQuoteList, rangeQuoteDetail, rangePrintInvItem, rangeQuoteDetailId As Range

    'branch variablbe
    Dim branchListRange As Range
    Dim lastRowBranch As Integer
    Dim branchName, branchId, bankName, accountName, accountNo, branchAddressKh, branchTel, branchEmail As String

    'setting variable
    Dim customerListRange,rng_file_setting As Range
    Dim folderProductPhoto, databasePath, databaseFile As String

    Set activeWorkbook = Workbooks("index.xlsm")
    Set ws_print_quote = activeWorkbook.Sheets("PrintQuotation")
    Set rangePrintInvItem = ws_print_quote.Range("C27:L143") 'Range("print_quote_items")

    If (Me.QuotationComboBox.Value <> "Please select quotation code" And Me.QuotationComboBox.Value <> "") Then 'checkDatabaseFile = TRUE And
        rangePrintInvItem.Value = "" 'rangePrintInvItem.ClearContents
        deletePicturesInPrintQuoteItem
        Set ws_file_setting     = activeWorkbook.Sheets("FilesSetting")
        databasePath            = ws_file_setting.Range("database_path").Value
        databaseFile            = databasePath & "database_kscm.xlsm"

        If (IsWorkbookOpen(databaseFile) = FALSE) Then
            ' Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        End If

        Set wbDatabase         = Workbooks("database_kscm.xlsm")
        Set ws_quote_data      = wbDatabase.Sheets("QuoteData")
        Set ws_quote_item_data = wbDatabase.Sheets("QuoteDetailData")
        Set ws_product         = activeWorkbook.Sheets("ProductList")

        lastQuoteDetailId      = ws_quote_item_data.Cells(Rows.Count, "C").End(xlUp).Row
        lastRowFileSetting     = ws_file_setting.Cells(Rows.Count,"C").End(xlUp).row

        Set rangeQuoteList     = ws_quote_data.Range("C8:AJ" & lastQuoteDetailId)    'Quote_list2
        Set rangeQuoteDetailId = ws_quote_item_data.Range("C8:C" & lastQuoteDetailId)
        Set rangeQuoteDetail   = ws_quote_item_data.Range("C8:Z" & lastQuoteDetailId)
        Set rng_file_setting   = ws_file_setting.Range("C8:M" & lastRowFileSetting)

        folderProductPhoto     = Application.Vlookup(40,rng_file_setting,10,FALSE)
        countQuoteDetailItem   = Application.CountA(rangeQuoteDetailId)
        quoteId                = CInt(Split(Me.QuotationComboBox.Value, "#")(1))
        quoteCode              = Split(Me.QuotationComboBox.Value, "#")(0)
        quoteCode              = Left(quoteCode, Len(quoteCode) - 1)

        customerId       = Application.VLookup(quoteId, rangeQuoteList, 6, False)
        quoteDate        = Format(Application.VLookup(quoteId, rangeQuoteList, 3, False), "yyyy/mm/dd")
        branchId         = Application.VLookup(quoteId, rangeQuoteList, 5, False)
        branchName       = ""
        customerName     = Application.VLookup(quoteId, rangeQuoteList, 7, False)
        quoteNote        = Application.VLookup(quoteId, rangeQuoteList, 24, False)

        vatPercent       = Application.VLookup(quoteId, rangeQuoteList, 12, False)
        vatCal           = Application.VLookup(quoteId, rangeQuoteList, 14, False)
        totalVat         = Application.VLookup(quoteId, rangeQuoteList, 18, False)

        exchangeRate     = Application.VLookup(quoteId, rangeQuoteList, 11, False)
        totalAmount      = Application.VLookup(quoteId, rangeQuoteList, 17, False)
        totalDiscount    = Application.VLookup(quoteId, rangeQuoteList, 19, False)
        totalDeposit     = Application.VLookup(quoteId, rangeQuoteList, 20, False)
        grandTotalAmount = Application.VLookup(quoteId, rangeQuoteList, 21, False)

        lastRowCustomer  = activeWorkbook.Sheets("CustomerList").Cells(Rows.Count, "C").End(xlUp).Row
        lastRowBranch    = activeWorkbook.Sheets("Branch").Cells(Rows.Count, "C").End(xlUp).Row
        lastRowUom       = activeWorkbook.Sheets("UoM").Cells(Rows.Count, "C").End(xlUp).Row
        lastProductRow   = ws_product.Cells(Rows.Count,"C").End(xlUp).Row

        Set customerListRange  = activeWorkbook.Sheets("CustomerList").Range("C5:P" & lastRowCustomer)
        Set branchListRange    = activeWorkbook.Sheets("Branch").Range("C7:AF" & lastRowBranch)
        Set rngUom             = activeWorkbook.Sheets("UoM").Range("C8:I" & lastRowUom)
        Set rngProductList     = ws_product.Range("C8:Z" & lastProductRow)

        customerNameEn         = Application.VLookup(customerId, customerListRange, 5, False)
        customerNameKh         = Application.VLookup(customerId, customerListRange, 6, False)
        customerAddressKh      = Application.VLookup(customerId, customerListRange, 9, False)
        customerTel            = Application.VLookup(customerId, customerListRange, 7, False)
        customerEmail          = Application.VLookup(customerId, customerListRange, 10, False)
        customerTinNumber      = Application.VLookup(customerId, customerListRange, 12, False)

        bankName               = Split(Application.VLookup(branchId, branchListRange, 6, False), ",")(0)
        accountName            = Split(Application.VLookup(branchId, branchListRange, 6, False), ",")(1)
        accountNo              = Split(Application.VLookup(branchId, branchListRange, 6, False), ",")(2)
        branchAddressKh        = Application.VLookup(branchId, branchListRange, 22, False)
        branchTel              = Application.VLookup(branchId, branchListRange, 18, False)
        branchEmail            = Application.VLookup(branchId, branchListRange, 20, False)

        'Quote Information
        ws_print_quote.Range("print_quote_code").Value          = quoteCode
        ws_print_quote.Range("print_quote_date").Value          = quoteDate
        ws_print_quote.Range("print_quote_note").Value          = quoteNote
        ws_print_quote.Range("print_quote_exchange_rate").Value = exchangeRate
    
        ' 'Amount Information
        ws_print_quote.Range("print_quote_sub_total").Value      = totalAmount
        ws_print_quote.Range("print_quote_toal_discount").Value  = totalDiscount
        ws_print_quote.Range("print_quote_toal_deposit").Value   = totalDeposit
        ws_print_quote.Range("print_quote_toal_vat").Value       = totalVat
        ws_print_quote.Range("print_quote_grand_total").Value    = grandTotalAmount
        ws_print_quote.Range("print_quote_grand_total_kh").Value = grandTotalAmount * exchangeRate

        ' 'Branch information
        ' ws_print_quote.Range("print_quote_bank_name").Value       = bankName
        ' ws_print_quote.Range("print_quote_account_name").Value    = accountName
        ' ws_print_quote.Range("print_quote_account_no").Value      = accountNo
        ws_print_quote.Range("print_quote_branch_address_kh").Value = branchAddressKh
        ws_print_quote.Range("print_quote_branch_telephone").Value  = branchTel
        ws_print_quote.Range("print_quote_branch_email").Value      = branchEmail
        
        ' 'Customer Information
        ' ws_print_quote.Range("print_quote_customer_name_en").Value    = customerNameEn
        ws_print_quote.Range("print_quote_customer_name_kh").Value      = customerNameKh
        ws_print_quote.Range("print_quote_customer_address_kh").Value   = customerAddressKh
        ws_print_quote.Range("print_quote_customer_tel").Value          = CStr(customerTel)
        ws_print_quote.Range("print_quote_customer_email").Value        = customerEmail
        ws_print_quote.Range("print_quote_customer_tin_number").Value   = customerTinNumber

        ws_print_quote.Range("C27" & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = False 'Show all row in range print Quote
        If quoteId > 0 And countQuoteDetailItem > 0 Then
            k = 0
            m = 0
            ' debug.Print "countQuoteDetailItem=" & countQuoteDetailItem
            For j = 1 To countQuoteDetailItem
                If (rangeQuoteDetail.Cells(j, 2) = quoteId) Then
                    ' disItem             = rangeQuoteDetail.Cells(j,6)
                    ' totalDisItem        = rangeQuoteDetail.Cells(j,6)
                    productCode = rangeQuoteDetail.Cells(j, 7)
                    productName = rangeQuoteDetail.Cells(j, 8)           'MID(Split(rangeQuoteDetail.Cells(j, 7),"#")(0), 7, 100)
                    qty         = rangeQuoteDetail.Cells(j, 9)
                    qtyFree     = rangeQuoteDetail.Cells(j, 10)
                    productUoM  = Application.VLookup(MID(rangeQuoteDetail.Cells(j, 11), 1, 3) * 1, rngUom, 5, False)
                    specProduct = Application.VLookup(rangeQuoteDetail.Cells(j, 6), rngProductList, 11, False)
                    unitPrice   = rangeQuoteDetail.Cells(j, 18)
                    totalPrice  = rangeQuoteDetail.Cells(j, 20)
                 
                    Dim criteriaRange1, criteriaRange2 As Range
                    Dim countItem As Integer

                    Set criteriaRange1 = ws_quote_item_data.Range("D8:D" & countQuoteDetailItem + 7) 'column Quote id
                    Set criteriaRange2 = ws_quote_item_data.Range("W8:W" & countQuoteDetailItem + 7) 'column status
                    countItem          = Application.COUNTIFS(criteriaRange1, quoteId, criteriaRange2, 1)

                    ' debug.Print "countItem=" & countItem
                    If (qty > 0) Then
                        k = k + 1
                        m = m + 1
                        ' If (countItem >= 26 And countItem <= 104) Then '26, 52, 78, 104
                        '     If (k Mod 26 = 0) Then
                        '         ' Debug.Print "***************k=" & k
                        '         k = k + 1
                        '     End If
                        ' Else
                        '     If (k Mod 13 = 0) Then ' k = 13 or k = 26 or k = 39 or k = 52 or k = 65
                        '         ' Debug.Print "***************k=" & k
                        '         k = k + 14
                        '     End If
                        ' End If

                        ' Debug.Print "k=" & k
                        rangePrintInvItem.Cells(k, 1) = m 'productCode
                        rangePrintInvItem.Cells(k, 2) = productName
                        rangePrintInvItem.Cells(k, 6) = qty
                        rangePrintInvItem.Cells(k, 7) = productUoM
                        rangePrintInvItem.Cells(k, 8) = unitPrice
                        rangePrintInvItem.Cells(k, 10)= totalPrice
                        rangePrintInvItem.Cells(k, 9) = totalPrice * exchangeRate

                        '********Insert Product Photo********************
                        k = k + 1
                        rangePrintInvItem.Cells(k, 2) = " - Item Code :  " & productCode & vbNewLine & " - Spec : " & specProduct 'Spec , item code, description
                        
                        ' Define the full path to your image
                        Dim imageName As String
                        imageName = FindImageFile(folderProductPhoto & productCode)
                        
                        If (imageName <> "") Then 
                            Dim targetCell As Range
                            Dim img As Shape

                            ' Define target pixel dimensions
                            Dim targetWidthPx As Double
                            Dim targetHeightPx As Double
                            'Resolution : 213x120 = (1920-1707)x(1080-960)
                            targetWidthPx  = 213  ' Set your desired width in pixels
                            targetHeightPx = 120  ' Set your desired height in pixels
                            
                            ' Convert pixels to points (1 pixel = 0.75 points)
                            Dim widthInPoints As Double
                            Dim heightInPoints As Double
                            widthInPoints  = targetWidthPx * 0.75
                            heightInPoints = targetHeightPx * 0.75

                            Set targetCell = rangePrintInvItem.Cells(k, 6) 'ws_print_quote.Range("B2") ' Change to your target cell
                            
                           ' Insert the picture at the coordinates of the target cell
                            Set img = ws_print_quote.Shapes.AddPicture( _
                                Filename:=imageName, _
                                LinkToFile:=msoFalse, _
                                SaveWithDocument:=msoTrue, _
                                Left:=targetCell.Left, _
                                Top:=targetCell.Top, _
                                Width:=-1, _
                                Height:=-1) ' Use -1 to keep original size initially

                                ' Optional: Resize and lock the picture to the cell dimensions
                            With img
                                .LockAspectRatio = msoFalse ' Set to msoTrue if you only want to force one dimension
                                .Width     = widthInPoints
                                .Height    = heightInPoints
                                .Left      = targetCell.Left + 70 * 0.75 '(targetCell.Width - .Width) / 2
                                .Top       = targetCell.Top + (targetCell.Height - .Height) / 2
                                ' Ensure the image moves and resizes if rows/columns change
                                .Placement = xlMoveAndSize 
                            End With
                        Else
                            rangePrintInvItem.Cells(k, 6) = ""
                        End If
                    End If
                End If
            Next j
            LastRow = k + 27
            ws_print_quote.Range("C" & LastRow & ":C" & rangePrintInvItem.Rows.Count + 26).EntireRow.Hidden = True 'Hide empty row
        End If
    Else
        rangePrintInvItem.Value = "" 'rangePrintInvItem.ClearContents
        deletePicturesInPrintQuoteItem
    End If
    ' OnEnd
End Sub

Sub deletePicturesInPrintQuoteItem()
    Dim activateWorkbook As Workbook
    Dim ws_quote As Worksheet
    Dim targetRange As Range
    Dim i As Long
    
    Set activateWorkbook = Workbooks("index.xlsm")
    Set ws_quote         = activateWorkbook.Sheets("PrintQuotation")
    Set targetRange      = ws_quote.Range("C27:L143") ' Define your worksheet and target range
    
    ' Loop backwards through shapes to avoid indexing issues during deletion
    For i = ws_quote.Shapes.Count To 1 Step -1
        With ws_quote.Shapes(i)
            ' Check if the shape is a picture (handles AddPicture and linked pictures)
            If .Type = msoPicture Or .Type = msoLinkedPicture Then
                ' Check if the picture's top-left cell falls inside the target range
                If Not Application.Intersect(.TopLeftCell, targetRange) Is Nothing Then
                    .Delete
                End If
            End If
        End With
    Next i
End Sub

Private Sub Worksheet_Activate()
    ' OnStart

    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_print_quote, ws_quote_data, ws_file_setting As Worksheet
    Dim rowQuoteCode, rowQuoteId As Range
    Dim i, k, CountaResultCell, lastQuoteId As Integer
    Dim databasePath, databaseFile As String

    Set activeWorkbook = Workbooks("index.xlsm")
    Set ws_print_quote = activeWorkbook.Sheets("PrintQuotation")
    
    ws_print_quote.QuotationComboBox.Clear
    ws_print_quote.QuotationComboBox.AddItem "Please select quotation code"

    If (checkDatabaseFile = True) Then
        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath = ws_file_setting.Range("database_path").Value
        databaseFile = databasePath & "database_kscm.xlsm"

        If (IsWorkbookOpen(databaseFile) = False) Then
            ' Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
        End If
        
        Set wbDatabase    = Workbooks("database_kscm.xlsm")
        Set ws_quote_data = wbDatabase.Sheets("QuoteData")
        lastQuoteId       = ws_quote_data.Cells(Rows.Count, "C").End(xlUp).Row - 7

        Set rowQuoteCode  = ws_quote_data.Range("D8:D" & ws_quote_data.Cells(Rows.Count, "C").End(xlUp).Row)
        Set rowQuoteId    = ws_quote_data.Range("C8:C" & ws_quote_data.Cells(Rows.Count, "C").End(xlUp).Row)

        For i = 1 To lastQuoteId
            If rowQuoteCode.Cells(i) <> "" Then
                ws_print_quote.QuotationComboBox.AddItem rowQuoteCode(i) & " #" & rowQuoteId(i)
            End If
        Next i
    Else
        ' MsgBox "Database file doesn't exist."
        ' Exit Sub
    End If

    ' OnEnd
End Sub

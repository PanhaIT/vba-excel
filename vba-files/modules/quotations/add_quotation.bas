Attribute VB_Name = "add_quotation"

Public Sub printquotLetterHead()
  ActiveSheet.Range("quot_print_letter_head").PrintPreview
End Sub

Public Sub printquotNoLetterHead()
  ActiveSheet.Range("quot_print_no_letter_head").PrintPreview
End Sub

Public Sub copyPrintLetterHead()
   'quot_copy_print_letter_head
   Sheets("AddQuotation").Range("quot_print_letter_head").Copy
   Sheets("AddQuotation").Range("quot_print_letter_head").select
End Sub

Public Sub copyPrintNoLetterHead()
   'quot_copy_print_no_letter_head
   ActiveSheet.Range("quot_print_no_letter_head").Copy
   ActiveSheet.Range("quot_print_no_letter_head").select
End Sub

Public Sub AddQuotation()
    OnStart
    Dim ws_quot_data, ws_add_quote,ws_quot As Worksheet
    Dim unitPrice, newUnitPrice,unitCostByUom,unitPriceBigUom, totalPrice, totalDeposit,totalDiscount,totalAmount,grandTotalAmount,unitCostBigUom,totalDisItem,disItem,totalVat As Double
    Dim quotStatusNote,customerId,quotSheet,branchId,locationId,priceTypeId,discountTypeId,saleId,saleName As string
    Dim customerName,salesquotNote,quotWeek,quotMonth,quotYear,branchName,msgSave As String
    Dim conversion, qtyOrder,qtyOrderSmall,quotStatus, quotNo, quotDetailId,itemType As Integer

    Dim productId, qty, qtyFree,convertion,smallValUom,vatPercent,vatSettingId,vatCal As Integer
    Dim productCode, productName, productUoM,quotCode, lotsNumber As String

    Dim item_row, quotDetailData, quotData, invTotal,ws_gld_data,ws_gl_data, invValuation As Range
    Dim i, j, a, b As Long
    Dim coaDefaultRange As Range
    Dim totalCost, totalPriceBeforDis,discountAmount As Double
    Dim quotId As Long
    Dim quotDate, createdDate,expiredDate,approveDate As Date
    Dim confirmBoxAlert As VbMsgBoxResult
    Dim approvedBy, dataInputPath,quotServiceType As String
    Dim isApprove,isClose,coaSaleDiscount,discountPercent As Integer

    msgSave           = "Are you want to save this quotation"
    confirmBoxAlert   = MsgBoxW(msgSave, vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Add Sales quot")
    quotNo            = ActiveSheet.Range("quot_no").Value
    quotCode          = ActiveSheet.Range("quot_code").Value
    customerId        = ActiveSheet.Range("quot_customer_id").Value
    quotDate          = ActiveSheet.Range("quot_date").Value 'Format(ActiveSheet.Range("quot_date").Value, "dd/mm/yyyy")
    branchId          = ActiveSheet.Range("quot_branch_id").Value
    quotServiceType   = ActiveSheet.Range("quot_service_type").Value 'quot_service_type_id = chart account id

    If (ActiveSheet.Range("quot_branch").Value <> "") Then 
        branchName        = MID(ActiveSheet.Range("quot_branch").Value,5,100)
    Else
        MsgBox "Branch is requied, try again."
        Exit Sub
    End If

    locationGroupId   = ActiveSheet.Range("quot_location_group_id").Value
    locationId        = ActiveSheet.Range("quot_location_id").Value
    priceTypeId       = ActiveSheet.Range("quot_price_type_id").Value
    discountTypeId    = ActiveSheet.Range("quot_discount_id").Value
    saleId            = ActiveSheet.Range("quot_seller_id").Value
    shipTo            = ActiveSheet.Range("quot_ship_to").value
    shipTel           = ActiveSheet.Range("quot_ship_telephone").value

    If (ActiveSheet.Range("quot_seller").Value <> "") Then 
        saleName      = MID(ActiveSheet.Range("quot_seller").Value,5,100)
    Else
        saleName      = "General"
    End If

    If (ActiveSheet.Range("quot_customer").Value <> "") Then 
        customerName      = MID(ActiveSheet.Range("quot_customer").Value,6,100)
    Else
        MsgBox "Customer is requied, try again."
        Exit Sub
    End If

    salesquotNote  = ActiveSheet.Range("quot_note").Value
    currencyId     = ActiveSheet.Range("quot_currency_id").Value

    If (ActiveSheet.Range("quot_date").Value > 0) Then 
        quotMonth  = month(ActiveSheet.Range("quot_date").Value)
        quotYear   = year(ActiveSheet.Range("quot_date").Value)
    Else
        MsgBox "quot date is requied, try again."
        Exit Sub
    End If

    quotWeek          = ActiveSheet.Range("quot_week").Value
    createdDate       = Format(Now(), "yyyy/mm/dd hh:mm:ss")

    vatPercent        = ActiveSheet.Range("quot_vat_percent").Value
    If (ActiveSheet.Range("quot_vat").Value <> "") Then 
        vatSettingId  = MID(ActiveSheet.Range("quot_vat").Value,1,2)*1
    Else
        vatSettingId  = 1
    End If
    vatCal            = ActiveSheet.Range("quot_vat_cal_id").Value
    totalVat          = Format(ActiveSheet.Range("quot_total_vat").Value,"0.00")

    totalDeposit      = Format(ActiveSheet.Range("quot_deposit").Value,"0.00")
    totalDiscount     = Format(ActiveSheet.Range("quot_discount").Value,"0.00")
    totalAmount       = Format(ActiveSheet.Range("quot_sub_total_amount").Value,"0.00")
    quotStatus        = Format(ActiveSheet.Range("quot_status").Value,"0.00")
    totalBalance      = Format(ActiveSheet.Range("quot_total_amount").Value,"0.00")

    isApprove            = 1 '0:approved,1:not approve
    isClose              = 0 '0:close,1:open
    quotStatus           = 1 '0:inactive,1:active
    approveDate          = quotDate
    approvedBy           = ActiveSheet.Range("quot_seller").Value

    If (discountTypeId=1) Then 'discount amount
        discountAmount   = Format(ActiveSheet.Range("quot_discount_amount").Value,"0.00")
        discountPercent = 0
    ElseIf (discountTypeId=2) Then 'discount percent
        discountAmount  = 0
        discountPercent  = Format(ActiveSheet.Range("quot_discount_percent").Value,"0.00")
    Else 'no discount
        discountAmount  = 0
        discountPercent = 0
    End If

    If (ActiveSheet.Range("quot_exchange_rate").Value > 0) Then 
        exchangeRate  = ActiveSheet.Range("quot_exchange_rate").Value
    Else
        exchangeRate  = 4000
    End If
    totalDeposit      = 0
    ' totalBalance      = totalAmount + totalVat - (totalDiscount + totalDeposit)

    Dim activeWorkbook, wbDatabase As Workbook
    Dim ws_product, ws_file_setting As Worksheet
    Dim databasePath, databaseFile As String
    Dim rngProductList As Range

    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_add_quote     = activeWorkbook.Sheets("AddQuotation")
    Set ws_product     = activeWorkbook.Sheets("ProductList")
    Set rngProductList = ws_product.Range("C8:AA" & ws_product.Cells(Rows.Count,"C").End(xlUp).Row)

    If confirmBoxAlert = vbYes And checkDatabaseFile = TRUE  Then

        Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
        databasePath        = ws_file_setting.Range("database_path").Value
        databaseFile        = databasePath & "database_kscm.xlsm"
        dataInputPath       = ws_file_setting.Range("data_input_path").value

        ' Dim check_access As Integer
        ' check_access = 0 
        If (IsWorkbookOpen(CStr(databaseFile)) = FALSE) Then
            Set wbDatabase = Workbooks.Open(databaseFile)  'add here the path of your text file
            ' check_access = 1
        End If

        Run resetAllFilter()
        quotStatus = quotStatus * 1

        If quotStatus > 0 Then
            If quotDate = "" Or priceTypeId = "" Or customerId = "" Or totalAmount = 0 Then
                MsgBox "Please Select require field, try again."
                Exit Sub
            Else
                Dim access As Integer
                access = 0
                If totalAmount > 0 Then

                    '***Start save quotation***
                    Set ws_quot_data     = wbDatabase.Sheets("QuoteData")
                    Set ws_quot_detail   = wbDatabase.Sheets("quoteDetailData")
                    Set quotData         = ws_quot_data.Range("C8:AJ8")
                    quotId               = ws_quot_data.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    Set quotDetailData   = ws_quot_detail.Range("C8:Z8")
                    quotDetailId         = ws_quot_detail.Cells(Rows.Count, "C").End(xlUp).Row - 6

                    Set item_row         = activeWorkbook.Sheets("AddQuotation").Range("I8:AM87")

                    ' debug.Print "quotId=" & quotId & " , quotDetailId=" & quotDetailId
                    ' Debug.Print "quotCode=" & quotCode & ", customerName=" & customerName

                    'insert quotation data
                    If (totalVat<=0) Then 
                        vatCal          = 0
                        vatSettingId    = 1
                        vatPercent      = 0
                    End If

                    quotData.Cells(quotId,1)   = quotId 'id
                    quotData.Cells(quotId,2)   = quotCode 'quot code
                    quotData.Cells(quotId,3)   = quotDate 'quot date
                    quotData.Cells(quotId,4)   = companyId 'company id
                    quotData.Cells(quotId,5)   = branchId 'branch id
                    quotData.Cells(quotId,6)   = customerId 'customer id
                    quotData.Cells(quotId,7)   = customerName 'customer name
                    quotData.Cells(quotId,8)   = quotServiceType 'service type
                    quotData.Cells(quotId,9)   = currencyId 'currency id
                    quotData.Cells(quotId,10)  = priceTypeId 'price type id
                    quotData.Cells(quotId,11)  = exchangeRate 'exchange rate

                    quotData.Cells(quotId,12)  = vatPercent 'vat percent
                    quotData.Cells(quotId,13)  = vatSettingId 'vat setting id
                    quotData.Cells(quotId,14)  = vatCal 'vat calculate before and after discount
                    quotData.Cells(quotId,15)  = discountAmount 'discount amount
                    quotData.Cells(quotId,16)  = discountPercent 'discount percent

                    quotData.Cells(quotId,17)  = totalAmount 'sub total
                    quotData.Cells(quotId,18)  = totalVat 'total vat
                    quotData.Cells(quotId,19)  = totalDiscount 'total discount
                    quotData.Cells(quotId,20)  = totalDeposit 'total deposit
                    quotData.Cells(quotId,21)  = totalBalance 'total amount

                    quotData.Cells(quotId,22)  = shipTo 'ship to
                    quotData.Cells(quotId,23)  = shipTel 'ship telephone
                    quotData.Cells(quotId,24)  = salesquotNote 'note
                    quotData.Cells(quotId,25)  = createdDate 'created
                    quotData.Cells(quotId,26)  = createdBy 'created by
                    quotData.Cells(quotId,27)  = approveDate 'approved
                    quotData.Cells(quotId,28)  = approvedBy 'approved by
                    quotData.Cells(quotId,29)  = isClose 'is close
                    quotData.Cells(quotId,30)  = isApprove 'is approve
                    quotData.Cells(quotId,31)  = quotStatus 'quotStatus
                    quotData.Cells(quotId,32)  = quotWeek 'week
                    quotData.Cells(quotId,33)  = quotMonth 'month
                    quotData.Cells(quotId,34)  = quotYear 'year

                    For a = 1 To item_row.Rows.Count
                        If item_row.Cells(a,1) <> "" Then 'old condition => WorksheetFunction.Count(rng.Rows(a)) <> 0

                            'insert quotation detail
                            conversion          = 1
                            productId           = item_row.Cells(a,1)
                            productCode         = item_row.Cells(a,3)
                            productName         = Application.VLookup(CInt(productId), rngProductList, 6, False) 'MID(Split(item_row.Cells(a,5),"#")(0),7,100)
                            qty                 = item_row.Cells(a, 6) * 1
                            qtyFree             = item_row.Cells(a, 7) * 1
                            productUoM          = item_row.Cells(a,8)
                            lotsNumber          = item_row.Cells(a,9)
                            expiredDate         = item_row.Cells(a,10)
                            newUnitPrice        = item_row.Cells(a,12)
                            unitPriceByUom      = item_row.Cells(a,13)
                            disItem             = Format(item_row.Cells(a,15),"0.00")
                            totalPrice          = Format(item_row.Cells(a,17),"0.00")
                            remark              = item_row.Cells(a,18)
                            If (item_row.Cells(a,19)>0) Then
                                conversion      = item_row.Cells(a,19)
                            End If
                            smallValUom         = item_row.Cells(a,20)
                            totalDisItem        = Format(item_row.Cells(a,21),"0.00")
                            unitCostByUom       = Format(item_row.Cells(a,22),"0.00")
                            ' unitCostBigUom      = Format(item_row.Cells(a,23),"0.00")
                            unitPriceBigUom     = Format(item_row.Cells(a,24),"0.00")
                            itemType            = item_row.Cells(a,25) '1 : product, 2:service

                            If (newUnitPrice > 0 ) Then 
                                unitPrice = Format(newUnitPrice, "0.00")
                            Else
                                unitPrice = Format(unitPriceByUom, "0.00")
                            End If
                            totalCost          = (qty + qtyFree) * unitCostByUom
                            totalPriceBeforDis = Format(qty * unitPrice, "0.00")
                            
                            qtyOrder        = (qty + qtyFree) / conversion '(qty + qtyFree) / smallValUom / convertion
                            qtyOrderSmall   = (qty + qtyFree) * (smallValUom / conversion)

                            quotDetailData.Cells(quotDetailId, 1)   = quotDetailId 'id
                            quotDetailData.Cells(quotDetailId, 2)   = quotId 'quot id
                            quotDetailData.Cells(quotDetailId, 3)   = quotCode 'quot number
                            quotDetailData.Cells(quotDetailId, 4)   = quotDate 'quot date

                            quotDetailData.Cells(quotDetailId, 5)   = customerId 'customer id
                            quotDetailData.Cells(quotDetailId, 6)   = productId 'product id
                            quotDetailData.Cells(quotDetailId, 7)   = productCode 'barcode product
                            quotDetailData.Cells(quotDetailId, 8)   = productName 'item description/service

                            quotDetailData.Cells(quotDetailId, 9)   = qty 'qty
                            quotDetailData.Cells(quotDetailId, 10)  = qtyFree 'qty free
                            
                            quotDetailData.Cells(quotDetailId, 11)  = productUoM 'uom
                            quotDetailData.Cells(quotDetailId, 12)  = productUoMId 'uom id
                            quotDetailData.Cells(quotDetailId, 13)  = conversion 'uom conversion
                            
                            quotDetailData.Cells(quotDetailId, 14)  = "" 'discount id
                            quotDetailData.Cells(quotDetailId, 15)  = disItem 'discount amount
                            quotDetailData.Cells(quotDetailId, 16)  = "" 'discount percent

                            quotDetailData.Cells(quotDetailId, 17)  = unitCostByUom 'unit cost
                            quotDetailData.Cells(quotDetailId, 18)  = unitPrice 'unit price
                            quotDetailData.Cells(quotDetailId, 19)  = disItem 'discount total price
                            quotDetailData.Cells(quotDetailId, 20)  = totalPrice 'total price

                            quotDetailData.Cells(quotDetailId, 21)  = 1 'quotStatus
                            quotDetailData.Cells(quotDetailId, 22)  = quotWeek 'week
                            quotDetailData.Cells(quotDetailId, 23)  = quotMonth 'month
                            quotDetailData.Cells(quotDetailId, 24)  = quotYear 'year
                            quotDetailId = quotDetailId + 1
                        End If
                    Next a
                    '*******End save quotation*******
                    access = 1
                End If

                '***********Reset field after save*****************
                activeWorkbook.Sheets("AddQuotation").Range("quot_column_product").Value = ""
                activeWorkbook.Sheets("AddQuotation").Range("quot_qty").Value = ""
                activeWorkbook.Sheets("AddQuotation").Range("quot_qty_free").Value =""
                activeWorkbook.Sheets("AddQuotation").Range("quot_column_uom").Value = ""
                activeWorkbook.Sheets("AddQuotation").Range("quot_column_new_unit_price").Value = ""
                activeWorkbook.Sheets("AddQuotation").Range("quot_discount_by_item").Value = ""
                activeWorkbook.Sheets("AddQuotation").Range("quot_note").Value = ""
                activeWorkbook.Sheets("AddQuotation").Range("quot_customer").value = activeWorkbook.Sheets("AddQuotation").Range("F11").value
                activeWorkbook.Sheets("AddQuotation").Range("quot_column_expired_date").Value = ""
                activeWorkbook.Sheets("AddQuotation").Range("quot_column_lote_number").Value = ""
                activeWorkbook.Sheets("AddQuotation").Range("quot_column_remark").Value = ""

                '**************Go to sheet quotation after save********************
                ws_quot_data.Select
                ws_quot_data.Range("C" & ws_quot_data.Cells(Rows.Count, "C").End(xlUp).Row & ":" & "AJ" & ws_quot_data.Cells(Rows.Count, "C").End(xlUp).Row).Select 'Select New insert quot row
                ' ws_quot_data.Rows(ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row).Select 'Select all row
                ' Run "PRint_Click"

                wbDatabase.Save '**************Save workbook database********************
                OnEnd
                MsgBox "Quotation " & quotCode & " saved successful."
                Exit Sub
            End If
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub todayPickerDate()
   ActiveSheet.Range("quot_exchange_rate").Value = Format(Now(), "mm/dd/yyyy")
End Sub

Private Sub getQuotExchangeRateNBC()
    ActiveSheet.Range("quot_exchange_rate").Value = Sheets("NBC_Exchange_Rate").Range("ExchangeRateToday").Value
End Sub

Private Sub addSeller()
   Run resetAllFilter()
   frmAddSeller.Show
End Sub

Private Sub addShipTo()
   Run resetAllFilter()
   frmAddShipTo.Show
End Sub

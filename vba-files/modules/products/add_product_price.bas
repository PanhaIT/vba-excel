Attribute VB_Name = "add_product_price"

Private Sub AddProductPriceByUom()
    OnStart
    Dim activeWorkbook As Workbook
    Dim ws_pro_price, ws_price_his, ws_product As Worksheet
    Dim i,j,k,m,n,p,q, productId, productUomId, branchId, priceTypeId, productPriceId, totalRowProductPrice,setType,checkExist As Integer
    Dim productCode, productName, messageBox, priceTypeLatinName As String
    Dim unitCost, lastSellingPrice, newAmount, oldUnitCost,oldAmountBefore, oldAmount As Double
    Dim rngProPriceInput, productPriceData, CountaRange,rngProPriceHis,rngProductList,rngPriceType As Range
    Dim createdDate As Date
    Dim uomConversionId, priceByUom, checkExistPrice, uomName As Variant
    Dim alertMsgBox As VbMsgBoxResult
    Dim productSmallValUom, lastProductRow, checkProductId,checkBranchId,checkPriceTypeId,checkUomId,checkSetType,checkIsActive As Integer

    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_pro_price     = activeWorkbook.Sheets("ProductPrice")
    Set ws_product       = activeWorkbook.Sheets("ProductList")
    Set ws_price_his     = activeWorkbook.Sheets("ProductPriceHistory")
    Set rngProPriceInput = ws_pro_price.Range("product_price_input_by_uom_list") 'product_price_input_by_uom_list=B10:W21
    Set CountaRange      = Range("B10:B21")
    Set productPriceData = ws_pro_price.Range("C28:T28")
    Set rngProPriceHis   = ws_price_his.Range("C8:P8")
    Set rngProductList   = ws_product.Range("C8:Z8")
    Set rngPriceType     = activeWorkbook.Sheets("Setting").Range("price_type_list")

    alertMsgBox          = MsgBoxW("Are you want to save this product price", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Add Product Price")

    totalRowProductPrice = ws_pro_price.Cells(Rows.Count, "C").End(xlUp).Row - 27
    productPriceId       = ws_pro_price.Cells(Rows.Count, "C").End(xlUp).Row - 26
    CountaResultCell     = Application.WorksheetFunction.Count(CountaRange)

    productId            = ActiveSheet.Range("product_price_pid").Value
    productCode          = ActiveSheet.Range("product_price_pcode").Value 'Split(ActiveSheet.Range("product_price_name").Value, "#")(1)
    productName          = ActiveSheet.Range("product_price_pname").Value 'MID(Split(ActiveSheet.Range("product_price_name").Value, "#")(0), 7, 100)
    productUomId         = ActiveSheet.Range("product_price_uom_id").Value
    productSmallValUom   = ActiveSheet.Range("product_price_small_val_uom").Value
    branchId             = ActiveSheet.Range("product_price_branch_id").Value
    createdDate          = Format(Now(), "yyyy/mm/dd hh:mm:ss")

    If (alertMsgBox = vbYes) Then
        For i = 1 To CountaResultCell
            If (rngProPriceInput.Cells(i, 3) <> "") Then 'Check if price type empty
                priceTypeId          = CInt(rngProPriceInput.Cells(i, 1))
                priceTypeName        = MID(rngProPriceInput.Cells(i, 3),4,100)
                unitCost             = rngProPriceInput.Cells(i, 5)
                lastSellingPrice     = rngProPriceInput.Cells(i, 6) 'last selling price
                totalUomConversion   = rngProPriceInput.Cells(i, 14)
                If (rngProPriceInput.Cells(i, 18) <> "") Then
                setType              = MID(rngProPriceInput.Cells(i, 18), 1, 1) * 1
                Else
                setType              = 1
                End If
                uomConversionId      = Split(rngProPriceInput.Cells(i, 15),";")
                uomName              = Split(rngProPriceInput.Cells(i, 16),";")
                priceByUom           = Split(rngProPriceInput.Cells(i, 17),";")
                checkExistPrice      = Split(rngProPriceInput.Cells(i, 22),";") '1:exist, 0:not exist
                newAmount            = 0
                isActive             = 1

                'Insert/update price by each uom
                k = 0

                For j = 1 To totalUomConversion
                    If (uomConversionId(k) * 1 > 0) Then
                        ' debug.Print "uomConversionId(k)=" & uomConversionId(k) & "priceByUom(k)=" & priceByUom(k)
                        'checkExistPrice=" & checkExistPrice(k)
                        If priceByUom(k) * 1 > 0 Then 'Insert only price bigger than zero
                            checkExist = checkAmountInProductPrice(productId * 1,branchId * 1,priceTypeId * 1,uomConversionId(k) * 1,setType * 1,isActive *1)

                            If (checkExist=1) Then 'checkExist = checkExistPrice(k)
                                '1= exist price => update price
                                For m = 1 To totalRowProductPrice
                                    If (productId = productPriceData.Cells(m, 2) AND branchId = productPriceData.Cells(m, 5) And priceTypeId = productPriceData.Cells(m, 6) And uomConversionId(k) * 1 =  productPriceData.Cells(m, 7) And productPriceData.Cells(m, 15) = setType And productPriceData.Cells(m, 18) = 1) Then
                                        If (productPriceData.Cells(m, 12) > 0) Then
                                            'Insert unit cost history
                                            n = ws_price_his.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                            rngProPriceHis.Cells(n, 1)  = n 'id
                                            rngProPriceHis.Cells(n, 2)  = productPriceData.Cells(m, 2) 'product id
                                            rngProPriceHis.Cells(n, 3)  = productPriceData.Cells(m, 3) 'product code
                                            rngProPriceHis.Cells(n, 4)  = productPriceData.Cells(m, 4) 'product name
                                            rngProPriceHis.Cells(n, 5)  = productPriceData.Cells(m, 5) 'branch id
                                            rngProPriceHis.Cells(n, 6)  = productPriceData.Cells(m, 6) 'price type id
                                            rngProPriceHis.Cells(n, 7)  = productPriceData.Cells(m, 7) 'uom id
                                            rngProPriceHis.Cells(n, 8)  = productPriceData.Cells(m, 8) 'Uom Name
                                            rngProPriceHis.Cells(n, 9)  = productPriceData.Cells(m, 9) 'Price Type Name
                                            rngProPriceHis.Cells(n, 10) = productPriceData.Cells(m, 10) 'old unit cost
                                            rngProPriceHis.Cells(n, 11) = productPriceData.Cells(m, 11) 'amount before
                                            rngProPriceHis.Cells(n, 12) = productPriceData.Cells(m, 12) 'amount
                                            rngProPriceHis.Cells(n, 13) = productPriceData.Cells(m, 13) 'percent
                                            rngProPriceHis.Cells(n, 14) = productPriceData.Cells(m, 14) 'add on
                                            rngProPriceHis.Cells(n, 15) = productPriceData.Cells(m, 15) 'set type
                                            rngProPriceHis.Cells(n, 16) = productPriceData.Cells(m, 16) 'created
                                            rngProPriceHis.Cells(n, 17) = productPriceData.Cells(m, 17) 'created by
                                            rngProPriceHis.Cells(n, 18) = productPriceData.Cells(m, 18) 'is active
                                            productPriceData.Cells(m, 12) = priceByUom(k) * 1 'Update exist price in product price list
                                        End If
                                    End If
                                    messageBox = "Product price is ready updated."
                                Next m
                            ElseIf (checkExist=0) Then 'checkExist = checkExistPrice(k)
                                ' 0= not exist price => insert new price in product price list
                                productPriceData.Cells(productPriceId, 1)  = productPriceId 'id
                                productPriceData.Cells(productPriceId, 2)  = productId 'product id
                                productPriceData.Cells(productPriceId, 3)  = productCode 'product code
                                productPriceData.Cells(productPriceId, 4)  = productName 'product name
                                productPriceData.Cells(productPriceId, 5)  = branchId 'branch id
                                productPriceData.Cells(productPriceId, 6)  = priceTypeId 'price type id
                                productPriceData.Cells(productPriceId, 7)  = uomConversionId(k) 'uom id
                                productPriceData.Cells(productPriceId, 8)  = uomName(k) 'Uom Name
                                productPriceData.Cells(productPriceId, 9)  = priceTypeName 'Price Type Name
                                productPriceData.Cells(productPriceId, 10) = unitCost 'old unit cost
                                productPriceData.Cells(productPriceId, 11) = lastSellingPrice 'amount before
                                productPriceData.Cells(productPriceId, 12) = priceByUom(k) * 1 'amount
                                productPriceData.Cells(productPriceId, 13) = 0 'percent
                                productPriceData.Cells(productPriceId, 14) = 0 'add on
                                productPriceData.Cells(productPriceId, 15) = 1 'set type
                                productPriceData.Cells(productPriceId, 16) = createdDate 'created
                                productPriceData.Cells(productPriceId, 17) = 1 'created by
                                productPriceData.Cells(productPriceId, 18) = 1 'is active
                                productPriceId = productPriceId + 1
                                messageBox = "New product price saved successful."
                            End If

                            'Update price by uom in product list
                            lastProductRow = ws_product.Cells(Rows.Count, "C").End(xlUp).Row - 7
                            For p = 1 To lastProductRow
                                If (productId = rngProductList.Cells(p, 1) And productUomId = CInt(uomConversionId(k))) Then 
                                    rngProductList.Cells(p, 15) = priceByUom(k) * 1 'Unit Price
                                End If
                            Next p
                        End If
                        k = k + 1
                    End If
                Next j
            End If
        Next i
        OnEnd
        MsgBox messageBox
    Else
        Exit Sub
    End If
End Sub

Private  Function checkAmountInProductPrice(productId As Integer,branchId As Integer, priceTypeId As Integer, uomId As Integer, setType As integer, isActive As Integer) As Integer
    OnStart
    Dim ws_pro_price As Worksheet
    Dim i,j,k,productPriceId As Integer
    Dim checkExistTmp As Integer
    Dim productPriceData As Range

    Set ws_pro_price     = Sheets("ProductPrice")
    Set productPriceData = Sheets("ProductPrice").Range("C28:T28")
    productPriceId       = ws_pro_price.Cells(Rows.Count, "C").End(xlUp).Row - 27

    If (productId >0  And priceTypeId >0 And uomId >0) Then
        For i = 1 To productPriceId
            checkProductId    = productPriceData.Cells(i, 2)
            checkBranchId     = productPriceData.Cells(i, 5)
            checkPriceTypeId  = productPriceData.Cells(i, 6)
            checkUomId        = productPriceData.Cells(i, 7)
            checkSetType      = productPriceData.Cells(i, 15)
            checkIsActive     = productPriceData.Cells(i, 18)
        
            If (productId = checkProductId AND branchId = checkBranchId And priceTypeId = checkPriceTypeId And uomId=checkUomId And setType = checkSetType And isActive = checkIsActive) Then
                checkExistTmp = 1
                Exit For
            Else
                checkExistTmp = 0
            End If
        Next i
    Else
        checkExistTmp = 0
    End If
    checkAmountInProductPrice = checkExistTmp
    OnEnd
End Function

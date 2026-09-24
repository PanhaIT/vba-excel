Attribute VB_Name = "import_product_price"

    '****************Import Product Price***********************
    ' 1 - No
    ' 2 - Photo
    ' 3 - Photo Path
    ' 4 - SKU
    ' 5 - Barcode
    ' 6 - Product Name
    ' 7 - Product Name Other
    ' 8 - Product Group
    ' 9 - Brand
    ' 10 - Spec
    ' 11 - Description
    ' 12 - Reorder Level
    ' 13 - Unit Cost
    ' 14 - Unit Price
    ' 15 - Stock Ending
    ' 16 - Small Val UoM
    ' 17 - Is Not For Sale
    ' 18 - Is Lots Number
    ' 19 - Is Expired Date
    ' 20 - Is Warranty
    ' 21 - type
    ' 22 - UoM
    ' 23 - Convertion 2
    ' 24 - Conversion 3
    ' 25 - Total Uom
    ' 26 - Array Uom Id
    ' 27 - Total Price Type
    ' 28 - តម្លៃលក់ដុំ ទូទៅ
    ' 29 - តម្លៃលក់រាយ ទូទៅ

Private Sub importProductPriceByUom()
    OnStart

    Dim activeWorkbook As Workbook
    Dim ws_import_pro, ws_price_his, ws_product,ws_uom_con,ws_pro_price As Worksheet
    Dim i,j,k,m,n,p,q, productId, productUomId, branchId, priceTypeId, ImportProductPriceId, totalRowProductImport,setType,checkExist As Integer
    Dim productCode, productName, messageBox, priceTypeLatinName,priceTypeName As String
    Dim unitCost, lastSellingPrice, newAmount, oldUnitCost,oldAmountBefore, oldAmount As Double
    Dim ImportProductPriceData,rngProPriceHis,rngProductList,rngPriceType As Range
    Dim createdDate As Date
    Dim uomConversionId, checkExistPrice, uomName, arrPriceType, arrUomId As Variant
    Dim alertMsgBox As VbMsgBoxResult
    Dim isActive, totalRowProductPrice, productPriceId, mainUomId, lastColumn, lastProductRow, checkProductId,checkBranchId,checkPriceTypeId,checkUomId,checkSetType,checkIsActive As Integer

    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_import_pro    = activeWorkbook.Sheets("ImportProductPrice")
    Set ws_pro_price     = activeWorkbook.Sheets("ProductPrice")
    Set ws_product       = activeWorkbook.Sheets("ProductList")
    Set ws_price_his     = activeWorkbook.Sheets("ProductPriceHistory")
    Set ws_uom_con       = activeWorkbook.Sheets("UomConversion")

    Set ImportProductPriceData = ws_import_pro.Range("C19:AG19") 'price AC19:AS19
    Set rngProPriceHis   = ws_price_his.Range("C8:P8")
    Set rngProductList   = ws_product.Range("C8:Z8")
    Set productPriceData = ws_pro_price.Range("C28:T28")
    Set rngPriceType     = activeWorkbook.Sheets("Setting").Range("price_type_list")

    arrPriceType         = Split(ws_import_pro.Range("import_pro_price_type_arr").Value,",")
    alertMsgBox          = MsgBoxW("Are you want to save these product prices", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Add Product Price")

    totalRowProductImport = ws_import_pro.Cells(Rows.Count, "C").End(xlUp).Row - 19
    totalRowProductPrice  = ws_pro_price.Cells(Rows.Count, "C").End(xlUp).Row - 27
    productPriceId        = ws_pro_price.Cells(Rows.Count, "C").End(xlUp).Row - 26
    branchId              = ActiveSheet.Range("import_pro_branch_id").Value
    createdDate           = Format(Now(), "yyyy/mm/dd hh:mm:ss")

    ' debug.Print "totalRowProductPrice=" totalRowProductPrice & " , productPriceId=" & productPriceId
    ' exit sub
    If ActiveSheet.Range("F2").Value = "" Then
        MsgBox "Please select branch name."
        Exit Sub
    End If

    ' debug.Print "totalRowProductImport=" & totalRowProductImport
    ' totalRowProductImport = 10
    If (alertMsgBox = vbYes) Then
        ' totalRowProductImport = 176
        lastColumn = 27 ' total colum from product no. to Total Price Type
        For i = 1 To totalRowProductImport
            productId            = ImportProductPriceData.Cells(i+1, 1)
            productCode          = ImportProductPriceData.Cells(i+1, 4)
            productBarcode       = ImportProductPriceData.Cells(i+1, 4)
            productName          = ImportProductPriceData.Cells(i+1, 6)
            productNameOther     = ImportProductPriceData.Cells(i+1, 7)
            productGroup         = ImportProductPriceData.Cells(i+1, 8)

            productBrand         = ImportProductPriceData.Cells(i+1, 9)
            productSpec          = ImportProductPriceData.Cells(i+1, 10)
            productDescription   = ImportProductPriceData.Cells(i+1, 11)
            productReorder       = ImportProductPriceData.Cells(i+1, 12)
            unitCost             = ImportProductPriceData.Cells(i+1, 13)
            unitPrice            = ImportProductPriceData.Cells(i+1, 14)

            endingStock          = ImportProductPriceData.Cells(i+1, 15)
            smallValueUom        = ImportProductPriceData.Cells(i+1, 16)
            isNotForSales        = ImportProductPriceData.Cells(i+1, 17)
            isLotsNumber         = ImportProductPriceData.Cells(i+1, 18)
            isExpiredDate        = ImportProductPriceData.Cells(i+1, 19)
            isWarranty           = ImportProductPriceData.Cells(i+1, 20)

            itemType             = ImportProductPriceData.Cells(i+1, 21)
            productUom           = ImportProductPriceData.Cells(i+1, 22)
            mainUomId            = CInt(MID(ImportProductPriceData.Cells(i+1, 22),1,3))
            totalUomConversion   = ImportProductPriceData.Cells(i+1, 25)
            arrUomId             = ImportProductPriceData.Cells(i+1, 26)
            productUomId         = CInt(MID(ImportProductPriceData.Cells(i+1, 22),1,3))
            totalPriceType       = ImportProductPriceData.Cells(i+1, 27)
            setType              = 1
            isActive             = 1
            m = 1 '176-179

            For j = 0 To totalPriceType - 1
                For k = 0 To totalUomConversion - 1
                    priceTypeId    = CInt(arrPriceType(j))
                    priceTypeName  = Application.VLookup(priceTypeId, rngPriceType, 3, False) 'MID(ImportProductPriceData.Cells(i, 3),4,100)

                    If (ImportProductPriceData.Cells(1, lastColumn + m) = priceTypeId) Then
                        If (totalUomConversion = 1) Then
                           productUomId   = CInt(MID(ImportProductPriceData.Cells(i+1, 22),1,3))
                        ElseIf (totalUomConversion > 1) Then
                            productUomId  = Split(ImportProductPriceData.Cells(i+1, 26),",")(k)
                        End If
                        
                        Dim conversion, countMiddleUom,countSmallUom, index,index1 As Integer
                        Dim criteriaRange, criteriaRange1, criteriaRange2, criteriaRange3,criteriaRange4 As Range
                        Dim middleValUom,middleUomId, priceByUom, priceBigUom As Double

                        index   = 5 'Small value uom
                        index1  = 2 'to uom id
                        Set criteriaRange  = ws_uom_con.Range("D8:J500") 'table uom conversion range
                        Set criteriaRange1 = ws_uom_con.Range("D8:D500") 'from uom id
                        Set criteriaRange2 = ws_uom_con.Range("E8:E500") 'to uom id
                        Set criteriaRange3 = ws_uom_con.Range("I8:I500") 'is small uom , 0=big uom, 1:small uom
                        Set criteriaRange4 = ws_uom_con.Range("J8:J500") 'is_active
                        'priceByUom = ImportProductPriceData.Cells(i+1, lastColumn + m) 

                        'Check main uom
                        If (productUomId = CInt(Split(ImportProductPriceData.Cells(i+1, 26),",")(k))) Then
                            conversion = 1
                        End If

                        'Check middle uom
                        countMiddleUom = Application.COUNTIFS(criteriaRange1, mainUomId, criteriaRange3, 0, criteriaRange4, 1)
                        If (countMiddleUom > 0) Then 
                            middleUomId    = Slookup(mainUomId * 1, ws_uom_con.Range("D8:J500"), CInt(index1), ws_uom_con.Range("I8:I500"), 0, ws_uom_con.Range("J8:J500"), 1)
                            middleValUom   = Slookup(mainUomId * 1, ws_uom_con.Range("D8:J500"), CInt(index), ws_uom_con.Range("I8:I500"), 0, ws_uom_con.Range("J8:J500"), 1)
                            If (CInt(middleUomId) = CInt(productUomId)) Then
                                conversion = middleValUom
                            End If
                        End If

                        'Check small uom
                        countSmallUom = Application.COUNTIFS(criteriaRange1, mainUomId, criteriaRange3, 1, criteriaRange4, 1)
                        If (countSmallUom > 0) Then 
                            smallUomId    = Slookup(mainUomId * 1, ws_uom_con.Range("D8:J500"), CInt(index1), ws_uom_con.Range("I8:I500"), 1, ws_uom_con.Range("J8:J500"), 1)
                            If (CInt(smallUomId) = CInt(productUomId)) Then
                                conversion = smallValueUom
                            End If
                        End If

                        If (conversion = 1 ) Then 
                            priceByUom = unitPrice
                            'priceByUom = ImportProductPriceData.Cells(i+1, lastColumn + m)  * (smallValueUom / conversion) 'pirce by uom and price type get from unit price
                        Else
                            priceByUom = ImportProductPriceData.Cells(i+1, lastColumn + m)  'ImportProductPriceData.Cells(i+1, lastColumn + m)  * (smallValueUom / conversion)
                        End If
                        
                        unitCost   = unitCost / conversion
                        ' debug.Print "smallValueUom=" & smallValueUom & " , conversion=" & conversion
                        ' debug.Print "productCode=" & productCode & ", unitCost = " & unitCost & "=> price = " & ImportProductPriceData.Cells(i+1, lastColumn + m) & " , priceByUom=" & priceByUom
                        
                        ' smallValueUom=10 , conversion=1
                        ' productCode=DKS0209, unitCost = 1.75=> price = 0.25 , priceByUom=0.25
                        ' smallValueUom=10 , conversion=10
                        ' productCode=DKS0209, unitCost = 0.175=> price = 0.25 , priceByUom=0.25

                        If (priceByUom >= 0) Then
                            checkExist = checkAmountInProductPrice(CInt(productId),CInt(branchId),CInt(priceTypeId),CInt(productUomId),CInt(setType),CInt(isActive))
                            ' Debug.Print "checkExist=" & checkExist & " , Unit Cost = " & unitCost & " ,unitPrice = " & unitPrice & " ,product code = " & productCode & " => price type id in product = " & priceTypeId & ", price by uom=" & priceByUom & " ,priceTypeName=" & priceTypeName & " ,productUomId=" & productUomId & " , conversion =" & conversion
                        
                            If (checkExist=1) Then 'checkExist = checkExistPrice(k)
                                '1= exist price => update price
                                ' For m = 1 To totalRowProductPrice
                                '     If (productId = productPriceData.Cells(m, 2) AND branchId = productPriceData.Cells(m, 5) And priceTypeId = productPriceData.Cells(m, 6) And uomConversionId(k) * 1 =  productPriceData.Cells(m, 7) And productPriceData.Cells(m, 15) = setType And productPriceData.Cells(m, 18) =1) Then
                                '         If (productPriceData.Cells(m, 12) > 0) Then
                                '             'Insert unit cost history
                                '             n = ws_price_his.Cells(Rows.Count, "C").End(xlUp).Row - 6
                                '             rngProPriceHis.Cells(n, 1)  = n 'id
                                '             rngProPriceHis.Cells(n, 2)  = productPriceData.Cells(m, 2) 'product id
                                '             rngProPriceHis.Cells(n, 3)  = productPriceData.Cells(m, 3) 'product code
                                '             rngProPriceHis.Cells(n, 4)  = productPriceData.Cells(m, 4) 'product name
                                '             rngProPriceHis.Cells(n, 5)  = productPriceData.Cells(m, 5) 'branch id
                                '             rngProPriceHis.Cells(n, 6)  = productPriceData.Cells(m, 6) 'price type id
                                '             rngProPriceHis.Cells(n, 7)  = productPriceData.Cells(m, 7) 'uom id
                                '             rngProPriceHis.Cells(n, 8)  = productPriceData.Cells(m, 8) 'Uom Name
                                '             rngProPriceHis.Cells(n, 9)  = productPriceData.Cells(m, 9) 'Price Type Name
                                '             rngProPriceHis.Cells(n, 10) = productPriceData.Cells(m, 10) 'old unit cost
                                '             rngProPriceHis.Cells(n, 11) = productPriceData.Cells(m, 11) 'amount before
                                '             rngProPriceHis.Cells(n, 12) = productPriceData.Cells(m, 12) 'amount
                                '             rngProPriceHis.Cells(n, 13) = productPriceData.Cells(m, 13) 'percent
                                '             rngProPriceHis.Cells(n, 14) = productPriceData.Cells(m, 14) 'add on
                                '             rngProPriceHis.Cells(n, 15) = productPriceData.Cells(m, 15) 'set type
                                '             rngProPriceHis.Cells(n, 16) = productPriceData.Cells(m, 16) 'created
                                '             rngProPriceHis.Cells(n, 17) = productPriceData.Cells(m, 17) 'created by
                                '             rngProPriceHis.Cells(n, 18) = productPriceData.Cells(m, 18) 'is active
                                '             productPriceData.Cells(m, 12) = priceByUom(k) 'Update exist price in product price list
                                '         End If
                                '     End If
                                '     messageBox = "Product price is ready updated."
                                ' Next m
                            ElseIf (checkExist=0) Then 'checkExist = checkExistPrice(k)
                                ' 0= not exist price => insert new price in product price list

                                ' productId            = ImportProductPriceData.Cells(i+1, 1)
                                ' productCode          = ImportProductPriceData.Cells(i+1, 4)
                                ' productBarcode       = ImportProductPriceData.Cells(i+1, 4)
                                ' productName          = ImportProductPriceData.Cells(i+1, 6)
                                ' productNameOther     = ImportProductPriceData.Cells(i+1, 7)
                                ' productGroup         = ImportProductPriceData.Cells(i+1, 8)

                                ' productBrand         = ImportProductPriceData.Cells(i+1, 9)
                                ' productSpec          = ImportProductPriceData.Cells(i+1, 10)
                                ' productDescription   = ImportProductPriceData.Cells(i+1, 11)
                                ' productReorder       = ImportProductPriceData.Cells(i+1, 12)
                                ' unitCost             = ImportProductPriceData.Cells(i+1, 13)
                                ' unitPrice            = ImportProductPriceData.Cells(i+1, 14)

                                ' endingStock          = ImportProductPriceData.Cells(i+1, 15)
                                ' smallValueUom        = ImportProductPriceData.Cells(i+1, 16)
                                ' isNotForSales        = ImportProductPriceData.Cells(i+1, 17)
                                ' isLotsNumber         = ImportProductPriceData.Cells(i+1, 18)
                                ' isExpiredDate        = ImportProductPriceData.Cells(i+1, 19)
                                ' isWarranty           = ImportProductPriceData.Cells(i+1, 20)

                                ' itemType             = ImportProductPriceData.Cells(i+1, 21)
                                ' productUom           = ImportProductPriceData.Cells(i+1, 22)
                                ' mainUomId            = CInt(MID(ImportProductPriceData.Cells(i+1, 22),1,3))
                                ' totalUomConversion   = ImportProductPriceData.Cells(i+1, 25)
                                ' arrUomId             = ImportProductPriceData.Cells(i+1, 26)
                                ' productUomId         = CInt(MID(ImportProductPriceData.Cells(i+1, 22),1,3))
                                ' totalPriceType       = ImportProductPriceData.Cells(i+1, 27)
                                
                                productPriceData.Cells(productPriceId, 1)  = productPriceId 'id
                                productPriceData.Cells(productPriceId, 2)  = productId 'product id
                                productPriceData.Cells(productPriceId, 3)  = productCode 'product code
                                productPriceData.Cells(productPriceId, 4)  = productName 'product name
                                productPriceData.Cells(productPriceId, 5)  = branchId 'branch id
                                productPriceData.Cells(productPriceId, 6)  = CInt(priceTypeId) 'price type id
                                productPriceData.Cells(productPriceId, 7)  = productUomId 'uom id
                                productPriceData.Cells(productPriceId, 8)  = "" 'Uom Name
                                productPriceData.Cells(productPriceId, 9)  = priceTypeName 'Price Type Name
                                productPriceData.Cells(productPriceId, 10) = unitCost 'old unit cost
                                productPriceData.Cells(productPriceId, 11) = priceByUom 'amount before
                                productPriceData.Cells(productPriceId, 12) = priceByUom 'amount
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
                            ' lastProductRow = ws_product.Cells(Rows.Count, "C").End(xlUp).Row - 7
                            ' For p = 1 To lastProductRow
                            '     If (productId = rngProductList.Cells(p, 1) And productUomId = CInt(uomConversionId(k))) Then 
                            '         rngProductList.Cells(p, 15) = CInt(priceByUom(k)) 'Unit Price
                            '     End If
                            ' Next p
                        End If
                    End If
                Next k
                m = m + 1
            Next j
            debug.Print "**********************************************************"
        Next i

        OnEnd
        MsgBox messageBox
    Else
        Exit Sub
    End If
End Sub

Private  Function checkAmountInProductPrice(productId As Integer,branchId As Integer, priceTypeId As Integer, uomId As Integer, setType As integer, isActive As Integer) As Integer
    Dim ws_import_pro As Worksheet
    Dim i,j,k,ImportProductPriceId As Integer
    Dim checkExistTmp As Integer
    Dim ImportProductPriceData As Range

    Set ws_import_pro          = Sheets("ImportProductPrice")
    Set ImportProductPriceData = Sheets("ImportProductPrice").Range("C28:T28")
    ImportProductPriceId       = ws_import_pro.Cells(Rows.Count, "C").End(xlUp).Row - 27

    If (productId >0  And priceTypeId >0 And uomId >0) Then
        For i = 1 To ImportProductPriceId
            checkProductId    = ImportProductPriceData.Cells(i, 2)
            checkBranchId     = ImportProductPriceData.Cells(i, 5)
            checkPriceTypeId  = ImportProductPriceData.Cells(i, 6)
            checkUomId        = ImportProductPriceData.Cells(i, 7)
            checkSetType      = ImportProductPriceData.Cells(i, 15)
            checkIsActive     = ImportProductPriceData.Cells(i, 18)
        
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
    checkAmountInImportProductPrice = checkExistTmp
End Function
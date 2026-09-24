VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet65"
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

Private Sub ProductAddPrice_Change()
    ' OnStart

    Dim activeWorkbook As Workbook
    Dim ws_pro_price, ws_product  As Worksheet
    Dim productId As Integer
    Dim productCode, productName As String
    Dim rngProductList As Range
 
    Set activeWorkbook = Workbooks("index.xlsm")
    Set ws_pro_price = activeWorkbook.Sheets("ProductPrice")
    Set ws_product = activeWorkbook.Sheets("ProductList")
 
    If (checkDatabaseFile = True And Me.ProductAddPrice.Value <> "" And Me.ProductAddPrice.Value <> "Please select product") Then
        ws_pro_price.Range("product_price_input_range").Value = ""
        ws_pro_price.Range("I10").Select
        productId = CInt(Split(Me.ProductAddPrice.Value, "#")(1))
        productCode = Split(Split(Me.ProductAddPrice.Value, "#")(0), " _ ")(0)
        productName = Split(Split(Me.ProductAddPrice.Value, "#")(0), " _ ")(1)

        Set rngProductList = ws_product.Range("product_list")
        ws_pro_price.Range("product_price_pid").Value = productId
        ws_pro_price.Range("product_price_uom").Value = Application.VLookup(productId, rngProductList, 8, False)
        ws_pro_price.Range("product_price_uom_id").Value = CInt(MID(Application.VLookup(productId, rngProductList, 8, False), 1, 3))
        ws_pro_price.Range("product_price_pcode").Value = productCode
        ws_pro_price.Range("product_price_pname").Value = productName
    Else
        ws_pro_price.Range("product_price_pid").Value = ""
        ws_pro_price.Range("product_price_pcode").Value = ""
        ws_pro_price.Range("product_price_pname").Value = ""
        ws_pro_price.Range("product_price_uom").Value = ""
        ws_pro_price.Range("product_price_uom_id").Value = ""
    End If

    ' OnEnd
End Sub

Private Sub Worksheet_Activate()
    ' OnStart

    Dim activeWorkbook As Workbook
    Dim ws_pro_price, ws_product As Worksheet
    Dim rowProductCode, rowProductId, rowProductName As Range
    Dim i, CountaResultCell, lastProductId As Integer

    Set activeWorkbook = Workbooks("index.xlsm")
    Set ws_pro_price = activeWorkbook.Sheets("ProductPrice")
    Set ws_product = activeWorkbook.Sheets("ProductList")
    lastProductId = ws_product.Cells(Rows.Count, "C").End(xlUp).Row - 7

    ws_pro_price.ProductAddPrice.Clear
    ws_pro_price.ProductAddPrice.AddItem "Please select product"
    ' ws_pro_price.Range("product_price_input_range").Value = ""
    ' ws_pro_price.Range("product_price_pid").Value = ""
    ' ws_pro_price.Range("product_price_pcode").Value = ""
    ' ws_pro_price.Range("product_price_pname").Value = ""
    ' ws_pro_price.Range("product_price_uom").Value = ""
    ' ws_pro_price.Range("product_price_uom_id").Value = ""

    Set rowProductId = ws_product.Range("C8:C" & ws_product.Cells(Rows.Count, "C").End(xlUp).Row)
    Set rowProductCode = ws_product.Range("F8:F" & ws_product.Cells(Rows.Count, "C").End(xlUp).Row)
    Set rowProductName = ws_product.Range("H8:H" & ws_product.Cells(Rows.Count, "C").End(xlUp).Row)

    For i = 1 To lastProductId
        If rowProductCode.Cells(i) <> "" Then
            ws_pro_price.ProductAddPrice.AddItem rowProductCode(i) & " _ " & rowProductName(i) & " #" & rowProductId(i)
        End If
    Next i
    
    ' OnEnd
End Sub



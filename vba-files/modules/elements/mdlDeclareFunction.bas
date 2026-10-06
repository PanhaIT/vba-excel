Attribute VB_Name = "mdlDeclareFunction"

Public Sub OnStart()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual 'xlCalculationManual , xlAutomatic
    Application.DisplayStatusBar = False
    
    ' Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
End Sub

Public Sub OnEnd()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    ' Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Sub

Public Function ConvertToInteger(myVar As String) As Integer
    Dim MyNumber As Integer
    MyNumber = 0
    
    If IsNumeric(myVar) Then
        MyNumber = CInt(myVar)
    End If
    
    ConvertToInteger = MyNumber
End Function

Public Function SpellNumber(ByVal MyNumber)
    Dim Dollars, Cents, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    ' String representation of amount.
    MyNumber = Trim(str(MyNumber))
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(MyNumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(MID(MyNumber, DecimalPlace + 1) & _
                    "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(right(MyNumber, 3))
        If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Dollars
        Case ""
            Dollars = "Zero US Dollars"
        Case "One"
            Dollars = "One US Dollar"
            Case Else
            Dollars = Dollars & " US Dollars"
    End Select
    Select Case Cents
        Case ""
            Cents = " and Zero Cents"
        Case "One"
            Cents = " and One Cent"
                Case Else
            Cents = " and " & Cents & " Cents"
    End Select
    SpellNumber = Dollars & Cents
End Function

Public Function SpellKhNumber(ByVal MyNumber)
    Dim Riel, Cents, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    ' String representation of amount.
    MyNumber = Trim(str(MyNumber))
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(MyNumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(MID(MyNumber, DecimalPlace + 1) & _
                    "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(right(MyNumber, 3))
        If Temp <> "" Then Riel = Temp & Place(Count) & Riel
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Riel
        Case ""
            Riel = " and Zero Riel"
        Case "One"
            Riel = " and One Riel"
        Case Else
            Riel = Riel & " Riel"
    End Select
    Select Case Cents
        Case ""
            Cents = " and Zero Cents"
        Case "One"
            Cents = " and One Cent"
        Case Else
            Cents = " and " & Cents & " Cents"
    End Select
    SpellKhNumber = Riel
End Function

' Converts a number from 100-999 into text
Public Function GetHundreds(ByVal MyNumber)
    Dim result As String
    If val(MyNumber) = 0 Then Exit Function
    MyNumber = right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If MID(MyNumber, 1, 1) <> "0" Then
        result = GetDigit(MID(MyNumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If MID(MyNumber, 2, 1) <> "0" Then
        result = result & GetTens(MID(MyNumber, 2))
    Else
        result = result & GetDigit(MID(MyNumber, 3))
    End If
    GetHundreds = result
End Function

' Converts a number from 10 to 99 into text.
Public Function GetTens(TensText)
    Dim result As String
    result = ""           ' Null out the temporary function value.
    If val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case val(TensText)
            Case 10: result = "Ten"
            Case 11: result = "Eleven"
            Case 12: result = "Twelve"
            Case 13: result = "Thirteen"
            Case 14: result = "Fourteen"
            Case 15: result = "Fifteen"
            Case 16: result = "Sixteen"
            Case 17: result = "Seventeen"
            Case 18: result = "Eighteen"
            Case 19: result = "Nineteen"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case val(Left(TensText, 1))
            Case 2: result = "Twenty "
            Case 3: result = "Thirty "
            Case 4: result = "Forty "
            Case 5: result = "Fifty "
            Case 6: result = "Sixty "
            Case 7: result = "Seventy "
            Case 8: result = "Eighty "
            Case 9: result = "Ninety "
            Case Else
        End Select
        result = result & GetDigit _
            (right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = result
End Function

' Converts a number from 1 to 9 into text.
Public Function GetDigit(Digit)
    Select Case val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

Public Function IsWorkbookOpen(wbName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo 0
    
    If Not wb Is Nothing Then IsWorkbookOpen = True
End Function

Public Function FindImageFile(ByVal baseFilePath As String) As String
    Dim extensions As Variant
    Dim ext As Variant
    Dim testPath As String
    
    ' Define the list of allowed image extensions
    extensions = Array(".png", ".jpg", ".jpeg")
    
    ' Loop through each extension to find a match
    For Each ext In extensions
        testPath = baseFilePath & ext
        
        ' Dir returns the filename if it exists, or "" if it does not
        If Dir(testPath) <> "" Then
            FindImageFile = testPath
            Exit Function ' Match found, exit early
        End If
    Next ext
    
    ' Return empty string if no file exists with these extensions
    FindImageFile = ""
End Function

Public Function checkFileIsExist(fileName As String, path As String) As Boolean
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(path)

    For Each oFile In oFolder.Files
        If (fileName = oFile.Name) Then
            checkFileIsExist = True
            ' Debug.Print "file =" & oFile.Name
            Exit For
        Else
            checkFileIsExist = False
        End If
    Next oFile
End Function

Public Function resetAllFilter()
    Dim ws, ws_file_setting As Worksheet
    Dim wb As Workbook
    Dim listObj As ListObject

    Set wb = Workbooks("database_kscm.xlsm")
    'Set wb = ActiveWorkbook
    'This is if you place the macro in your personal wb to be able to reset the filters on any wb you're currently working on. Remove the set wb = thisworkbook if that's what you need
    For Each ws In wb.Worksheets
        If ws.FilterMode Then
            ws.ShowAllData
        Else

        End If
        'This removes "normal" filters in the workbook - however, it doesn't remove table filters
        For Each listObj In ws.ListObjects
            If listObj.ShowHeaders Then
                listObj.AutoFilter.ShowAllData
                listObj.Sort.SortFields.Clear
            End If
        Next listObj
    Next
    'And this removes table filters. You need both aspects to make it work.
End Function

Public Function resetWorkSheetsFilter()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim listObj As ListObject
    Set wb = ThisWorkbook
    'Set wb = ActiveWorkbook
    'This is if you place the macro in your personal wb to be able to reset the filters on any wb you're currently working on. Remove the set wb = thisworkbook if that's what you need
    For Each ws In wb.Worksheets
        If ws.FilterMode Then
            ws.ShowAllData
        Else

        End If
        'This removes "normal" filters in the workbook - however, it doesn't remove table filters
        For Each listObj In ws.ListObjects
            If listObj.ShowHeaders Then
                listObj.AutoFilter.ShowAllData
                listObj.Sort.SortFields.Clear
            End If
        Next listObj
    Next
    'And this removes table filters. You need both aspects to make it work.
End Function

Public Function PadStr(Expression As Variant, length As Integer, Optional padChar As String = " ", Optional alignment As XlHAlign = xlHAlignGeneral) As String
    'Pads a string with a given character.
    '@Expression - the string to pad
    '@length - the minimum length of the string (if @Expression is longer than @length, the original Expression will be returned)
    '@padChar - the character to pad with (a space by default)
    '@alignment - what type of alignment to use. Uses the XlAlign object for enumeration.
    '   xlHAlignLeft            - (Default) Aligns input text to the left
    '   xlHAlignGeneral         - Same as Default
    '   xlHAlignRight           - Aligns input text to the right
    '   xlHAlignCenter          - Center aligns text
    '   xlHAlignCenterAcrossSelection       - Same as xlHAlignCenter
    '   xlHAlignDistributed     - Distributes the text evenly within the length specified
    '   xlHAlignJustify         - Same as xlHAlignDistributed
    '   xlHAlignFill            - Fills the specified length with the text
    'example: if input is "ABC", " ", "8", see code below for what the output will be given the different direction options
    Dim direction As String
    If Len(Expression) >= length Or (padChar = "" And alignment <> xlHAlignFill) Then
        'if input is longer than pad-length padChar or no input given for padChar (note: padChar doesn't matter when
        'using xlHAlignFill) just return the input
        PadStr = Expression
    ElseIf Len(padChar) <> 1 And alignment <> xlHAlignFill Then
        'give error if padChar is not exactly 1 char in length (again, padChar doesn't matter when using xlHAlignFill)
        'padChar must be 1 char long because string() only accepts 1 char long input.
        Err.Raise vbObjectError + 513, , "input:'padChar' must have length 1." & vbNewLine & "SUB:PadStr"
    Else
        Dim pStr As String, i As Long
        Select Case alignment
            Case xlHAlignLeft, xlHAlignGeneral '(Default)
                '"ABC     "
                PadStr = CStr(Expression) & String(length - Len(CStr(Expression)), padChar)
            Case xlHAlignRight
                '"     ABC"
                PadStr = String(length - Len(CStr(Expression)), padChar) & CStr(Expression)
            Case xlHAlignCenter, xlHAlignCenterAcrossSelection
                '"   ABC  "
                pStr = String(Application.WorksheetFunction.RoundUp((length / 2) - (Len(Expression) / 2), 0), padChar)
                PadStr = pStr & Expression & pStr
            Case xlHAlignDistributed, xlHAlignJustify
                '"  A B C "       ("  A  B C " if lenth=9)
                Dim insPos As Long, loopCntr As Long: loopCntr = 1
                PadStr = Expression
                Do While Len(PadStr) < length
                    For i = 1 To Len(Expression)
                        PadStr = Left(PadStr, insPos) & padChar & right(PadStr, Len(PadStr) - insPos)
                        insPos = insPos + 1 + loopCntr
                        If Len(PadStr) >= length Then Exit For
                    Next i
                    PadStr = PadStr & padChar
                    loopCntr = loopCntr + 1
                    insPos = 0
                Loop
            Case xlHAlignFill
                '"ABCABCAB"
                For i = 1 To Application.WorksheetFunction.RoundUp(length / Len(Expression), 0)
                    PadStr = PadStr & Expression
                Next i
            Case Else
                'error
                direction = ""
                Err.Raise vbObjectError + 513, , "PadStr does not support the direction input ( " & direction & ")." & vbNewLine & "SUB:PadStr"
        End Select
        PadStr = Left(PadStr, length) 'output cannot be longer than the given length
    End If
End Function

Public Function checkDatabaseFile() As Boolean
    Dim activeWorkbook As Workbook
    Dim ws_file_setting As Worksheet
    Dim MyFSO As New FileSystemObject
    Dim databasePath As String
    Dim databaseFile, checkExistDb As Boolean
    
    Set activeWorkbook = Workbooks("index.xlsm")
    Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
    databasePath = ws_file_setting.Range("database_path").Value

    databaseFile = MyFSO.FileExists(databasePath & "database_kscm.xlsm")
    If (databaseFile = True) Then
        checkExistDb = True
    Else
        checkExistDb = False
    End If
    checkDatabaseFile = checkExistDb
End Function

Public  Function checkFolderExist()
    Dim activeWorkbook As Workbook
    Dim ws_file_setting As Worksheet
    Dim rng_file_setting As Range
    Dim lastRowFileSetting As Integer
    Dim MyFSO As New FileSystemObject
    Dim Pth As String
    Dim i As Integer
    Dim fullPath, mainPath, folder,subFolder1,subFolder2,subFolder3,subFolder4 As String
    
    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_file_setting  = activeWorkbook.Sheets("FilesSetting")
    lastRowFileSetting   = ws_file_setting.Cells(Rows.Count,"C").End(xlUp).row
    Set rng_file_setting = ws_file_setting.Range("C8:M" & lastRowFileSetting)

    folder     = ""
    subFolder1 = ""
    subFolder2 = ""
    subFolder3 = ""
    subFolder4 = ""
    For i = 1 To lastRowFileSetting - 7
        mainPath = rng_file_setting.Cells(i,4)
        If (mainPath <> "") Then
            If (rng_file_setting.Cells(i,5) <> "") Then
                folder = mainPath & rng_file_setting.Cells(i,5) 'ks_system
                If MyFSO.FolderExists(folder) = False Then MyFSO.CreateFolder (folder)
            End If
            If (MyFSO.FolderExists(folder) = True) Then
                If (rng_file_setting.Cells(i,6) <> "") Then
                    subFolder1 = folder & "\" & rng_file_setting.Cells(i,6) 'report
                    If MyFSO.FolderExists(subFolder1) = False Then MyFSO.CreateFolder (subFolder1)
                End If
            End If
            If (MyFSO.FolderExists(subFolder1) = True) Then
                If (rng_file_setting.Cells(i,7) <> "") Then
                    subFolder2 = subFolder1 & "\" & rng_file_setting.Cells(i,7) 'subFolder2
                    If MyFSO.FolderExists(subFolder2) = False Then MyFSO.CreateFolder (subFolder2)
                End If
            End If
            If (MyFSO.FolderExists(subFolder2) = True) Then
                If (rng_file_setting.Cells(i,8) <> "") Then
                    subFolder3 = subFolder2 & "\" & rng_file_setting.Cells(i,8) 'sub folder 3
                    If MyFSO.FolderExists(subFolder3) = False Then MyFSO.CreateFolder (subFolder3)
                End If
            End If
            If (MyFSO.FolderExists(subFolder3) = True) Then
                If (rng_file_setting.Cells(i,9) <> "") Then
                    subFolder4 = subFolder3 & "\" & rng_file_setting.Cells(i,9) 'sub folder 4
                    If MyFSO.FolderExists(subFolder4) = False Then MyFSO.CreateFolder (subFolder4)
                End If
            End If
        End If
    Next i
End Function

Public Function getStockAvariableByUom(productId As Integer, smallQtyAvariable As Long, smallValUom As Integer, status As Integer, mainUomId As Integer, Optional mainUomName As String) As String
    OnStart
    Dim activeWorkbook As Workbook
    Dim ws_uom,ws_uom_con As Worksheet
    Dim rng_uom_list As Range

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_uom          = activeWorkbook.Sheets("UoM")
    Set rng_uom_list    = ws_uom.Range("C8:I" & ws_uom.Cells(Rows.Count,"C").End(xlUp).row)

    If (mainUomName = "") Then mainUomName = Application.VLOOKUP(mainUomId,rng_uom_list,5,FALSE)

    If (smallValUom = 1) Then 
        getStockAvariableByUom = smallQtyAvariable & mainUomName
    Else
        Dim totalConversion, smallmainUomId, middlemainUomId, middleValUom, countMiddleUom, countSmallUom, index ,index1 As Integer
        Dim criteriaRange, criteriaRange1, criteriaRange2, criteriaRange3,criteriaRange4 As Range
        Dim smallUomName,middleUomName As String
        Dim bigQty, decimalQty,qtyNegativeAndPositive As Double
        Dim integerQty As Long

        Set ws_uom_con  = activeWorkbook.Sheets("UomConversion")

        index   = 5 'value uom
        index1  = 2 'to uom id
        Set criteriaRange  = ws_uom_con.Range("D8:J500") 'table uom conversion range
        Set criteriaRange1 = ws_uom_con.Range("D8:D500") 'from uom id
        Set criteriaRange2 = ws_uom_con.Range("E8:E500") 'to uom id
        Set criteriaRange3 = ws_uom_con.Range("I8:I500") 'is small uom , 0=big uom, 1:small uom
        Set criteriaRange4 = ws_uom_con.Range("J8:J500") 'is_active

        totalConversion = Application.COUNTIFS(criteriaRange1, mainUomId, criteriaRange4, 1) 'Count uom conversion base on main uom id
        
        'Check small uom
        countSmallUom = Application.COUNTIFS(criteriaRange1, mainUomId, criteriaRange3, 1, criteriaRange4, 1)
        If (countSmallUom > 0) Then 
            smallmainUomId  = Slookup(mainUomId * 1, ws_uom_con.Range("D8:J500"), CInt(index1), ws_uom_con.Range("I8:I500"), 1, ws_uom_con.Range("J8:J500"), 1)
            smallUomName    = Application.VLOOKUP(smallmainUomId,rng_uom_list,5,FALSE)
        End If

        bigQty      = smallQtyAvariable/smallValUom
        integerQty  = Fix(smallQtyAvariable/smallValUom)
        decimalQty  = smallQtyAvariable/smallValUom - Fix(integerQty)

        '******product have 2 uoms => totalConversion = 1
        If (totalConversion = 1) Then 
            If (smallQtyAvariable Mod smallValUom) = 0 Then
                getStockAvariableByUom = bigQty & mainUomName
            Else
                decimalQty = CInt(decimalQty * smallValUom)
                If (integerQty = 0) Then getStockAvariableByUom = decimalQty & smallUomName Else getStockAvariableByUom = integerQty & mainUomName & " & " & decimalQty & smallUomName
            End If
        Else
            '******product have 3 uoms => totalConversion = 2
            'Check middle uom
            countMiddleUom = Application.COUNTIFS(criteriaRange1, mainUomId, criteriaRange3, 0, criteriaRange4, 1)
            If (countMiddleUom > 0) Then 
                middlemainUomId  = Slookup(mainUomId * 1, ws_uom_con.Range("D8:J500"), CInt(index1), ws_uom_con.Range("I8:I500"), 0, ws_uom_con.Range("J8:J500"), 1)
                middleValUom     = Slookup(mainUomId * 1, ws_uom_con.Range("D8:J500"), CInt(index), ws_uom_con.Range("I8:I500"), 0, ws_uom_con.Range("J8:J500"), 1)
                middleUomName    = Application.VLOOKUP(middlemainUomId,rng_uom_list,5,FALSE)
            End If

            If (smallQtyAvariable Mod smallValUom) = 0 Then
                getStockAvariableByUom = bigQty & mainUomName
            Else
                decimalQty = decimalQty * middleValUom
                If decimalQty = Int(decimalQty) Then
                    If (integerQty = 0) Then getStockAvariableByUom = decimalQty & middleUomName Else getStockAvariableByUom = integerQty & mainUomName & " & " & decimalQty & middleUomName
                Else
                    If (decimalQty < 0) Then  qtyNegativeAndPositive = Int(decimalQty * -1) * -1 Else qtyNegativeAndPositive = Int(decimalQty)
                    smallestQty = (decimalQty-qtyNegativeAndPositive)*smallValUom/middleValUom

                    If (integerQty = 0) Then
                        If (qtyNegativeAndPositive = 0) Then getStockAvariableByUom  =  smallestQty & smallUomName Else getStockAvariableByUom = qtyNegativeAndPositive & middleUomName & " & " & smallestQty & smallUomName
                    Else
                        If (qtyNegativeAndPositive = 0) Then getStockAvariableByUom = integerQty & mainUomName & " & " & smallestQty & smallUomName Else  getStockAvariableByUom = integerQty & mainUomName & " & " & qtyNegativeAndPositive & middleUomName & " & " & smallestQty & smallUomName
                    End If
                End If
            End If
        End If
    End If
    OnEnd
End Function

Public  Sub getItemInvoiceTest()
    OnStart
    Dim activeWorkbook As Workbook
    Dim wsInv As Worksheet
    Dim lastRow As Long
    Dim item_row,rngPro As Range
    Dim i As Long
    Dim a,incNo As Integer

    Set activeWorkbook = Workbooks("index.xlsm")
    Set wsInv = activeWorkbook.Sheets("ProductList")
    
    lastRow = wsInv.Cells(Rows.Count, "C").End(xlUp).Row
    Set item_row  = wsInv.Range("C8:AA" & lastRow) 'product_list ,"C8:AA" & lastRow
    incNo=1
    For a = 1 To item_row.rows.Count
        If (item_row.Cells(a,1) <> "") Then
            debug.Print item_row.Cells(a,4)  '"i=" & i & ", lastRow=" & lastRow & ", sku=" & item_row.Cells(i, 4)\
            wsInv.Cells(a + 7,"AI").Value = item_row.Cells(a,4)
        End If
    Next a

    MsgBox "Invoice populated successfully!"
    OnEnd
End Sub

' AI code VBA excel sheet get auto select items when create invoice by passing argument total amount of invoice, multi product group,multi product brand,mu
Function AutoSelectInvoiceItems(ByVal TargetAmount As Double, _
                                ByVal ProductGroups As String, _
                                ByVal ProductBrands As String, _
                                ByVal Locations As String, _
                                ByVal Warehouses As String, _
                                ByVal Branch As String) As Variant
    
    Dim wsInv As Worksheet
    Dim lastRow As Long, i As Long
    Dim currentAmount As Double
    Dim itemPrice As Double, itemQty As Double, lineTotal As Double
    
    ' Output array to hold selected item details
    ' Adjust size or columns based on what you need to return
    Dim selectedItems() As Variant
    Dim itemCount As Long
    itemCount = 0
    currentAmount = 0#
    
    ' 1. Set your source ProductList sheet
    On Error Resume Next
    Set wsInv = ThisWorkbook.Sheets("ProductList")
    On Error GoTo 0
    
    If wsInv Is Nothing Then
        MsgBox "Error: 'ProductList' sheet not found!", vbCritical
        Exit Function
    End If
    
    lastRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row
    
    ' 2. Loop through ProductList items (Assuming headers are in row 1, data starts row 2)
    ' Adjust column letters (e.g., "A", "B", "C") to match your exact ProductList sheet layout
    For i = 2 To lastRow
        
        ' Check if we have already reached or exceeded the required invoice amount
        If currentAmount >= TargetAmount Then Exit For
        
        ' Extract criteria values from the current row
        Dim rowGroup As String: rowGroup = wsInv.Cells(i, "A").Value      ' Column A: Product Group
        Dim rowBrand As String: rowBrand = wsInv.Cells(i, "B").Value      ' Column B: Brand
        Dim rowLoc As String: rowLoc = wsInv.Cells(i, "C").Value          ' Column C: Location
        Dim rowWh As String: rowWh = wsInv.Cells(i, "D").Value            ' Column D: Warehouse
        Dim rowBranch As String: rowBranch = wsInv.Cells(i, "E").Value    ' Column E: Branch
        
        ' 3. Apply Multi-Filter Checks (using InStr to handle comma-separated multi-selects)
        If (ProductGroups = "" Or InStr(1, ProductGroups, rowGroup, vbTextCompare) > 0) And _
           (ProductBrands = "" Or InStr(1, ProductBrands, rowBrand, vbTextCompare) > 0) And _
           (Locations = "" Or InStr(1, Locations, rowLoc, vbTextCompare) > 0) And _
           (Warehouses = "" Or InStr(1, Warehouses, rowWh, vbTextCompare) > 0) And _
           (Branch = "" Or LCase(rowBranch) = LCase(Branch)) Then

            ' Extract item financial/quantity data
            Dim itemID As String: itemID = wsInv.Cells(i, "F").Value      ' Column F: Item ID/SKU
            itemQty = wsInv.Cells(i, "G").Value                           ' Column G: Available Qty
            itemPrice = wsInv.Cells(i, "H").Value                         ' Column H: Unit Price
            
            If itemQty > 0 And itemPrice > 0 Then
                lineTotal = itemQty * itemPrice
                
                ' Check if adding the whole lot exceeds the target amount
                If (currentAmount + lineTotal) > TargetAmount Then
                    ' Calculate exactly how many pieces are needed to hit the target
                    Dim neededAmount As Double
                    neededAmount = TargetAmount - currentAmount
                    
                    Dim neededQty As Double
                    neededQty = Application.WorksheetFunction.RoundUp(neededAmount / itemPrice, 0)
                    
                    ' If available stock covers the needed partial quantity
                    If neededQty <= itemQty Then
                        itemQty = neededQty
                        lineTotal = itemQty * itemPrice
                    End If
                End If
                
                ' Update total invoice accumulator
                currentAmount = currentAmount + lineTotal
                itemCount = itemCount + 1
                
                ' Resize array and store the selected item data
                ReDim Preserve selectedItems(1 To 4, 1 To itemCount)
                selectedItems(1, itemCount) = itemID      ' SKU
                selectedItems(2, itemCount) = itemQty     ' Quantity to pull
                selectedItems(3, itemCount) = itemPrice   ' Price
                selectedItems(4, itemCount) = lineTotal   ' Total for line
            End If
            
        End If
    Next i
    
    ' Return the populated array back to the calling sub
    If itemCount > 0 Then
        AutoSelectInvoiceItems = selectedItems
    Else
        AutoSelectInvoiceItems = Empty
    End If
    
End Function

Sub addItemsInvoice()
    Dim invoiceData As Variant
    Dim wsInvoice As Worksheet
    Dim targetAmt As Double
    Dim i As Long
    
    Set wsInvoice = ThisWorkbook.Sheets("Invoice")
    targetAmt = 5000.00 ' Your target invoice total amount
    
    ' Pass comma-separated strings for multi-select arguments
    invoiceData = AutoSelectInvoiceItems(targetAmt, _
                                         "Electronics,Appliances", _
                                         "Sony,Samsung", _
                                         "North,East", _
                                         "WH-01,WH-02", _
                                         "Main Branch")
                                         
    ' Check if items were found
    If IsEmpty(invoiceData) Then
        MsgBox "No items matched the criteria or stock is empty.", vbExclamation
        Exit Sub
    End If
    
    ' Clear old invoice lines (assuming rows 5 onwards are item rows)
    wsInvoice.Rows("5:100").ClearContents
    
    ' Write the array data down onto the invoice template
    For i = 1 To UBound(invoiceData, 2)
        wsInvoice.Cells(4 + i, "A").Value = invoiceData(1, i) ' Item ID
        wsInvoice.Cells(4 + i, "B").Value = invoiceData(2, i) ' Qty Allocated
        wsInvoice.Cells(4 + i, "C").Value = invoiceData(3, i) ' Price
        wsInvoice.Cells(4 + i, "D").Value = invoiceData(4, i) ' Total
    Next i
    
    MsgBox "Invoice populated successfully!", vbInformation
End Sub

' Public Sub CreateInvoiceByCriteria(ByVal TargetAmount As Double, _
'                             ByVal ProdGroup As String, _
'                             ByVal ProdBrand As String, _
'                             ByVal Location As String, _
'                             ByVal Warehouse As String, _
'                             ByVal Branch As String)

'     Dim wsInventory As Worksheet
'     Dim wsInvoice As Worksheet
'     Dim lastRowInv As Long
'     Dim nextRowInvc As Long
'     Dim i As Long
    
'     Dim currentGroup As String
'     Dim currentBrand As String
'     Dim currentLoc As String
'     Dim currentWh As String
'     Dim currentBranch As String
'     Dim itemPrice As Double
'     Dim itemName As String
'     Dim itemId As String
    
'     Dim RunningTotal As Double
'     Dim MatchCount As Long
    
'     ' Set worksheet references (Adjust names to match your workbook)
'     Set wsInventory = ThisWorkbook.Sheets("Inventory")
'     Set wsInvoice = ThisWorkbook.Sheets("Invoice")
    
'     ' Find the last row of data in the Inventory sheet (assuming Column A has Item IDs)
'     lastRowInv = wsInventory.Cells(wsInventory.Rows.Count, "A").End(xlUp).Row
    
'     ' Find the next available row in the Invoice sheet (assuming Column A)
'     nextRowInvc = wsInvoice.Cells(wsInvoice.Rows.Count, "A").End(xlUp).Row + 1
    
'     RunningTotal = 0
'     MatchCount = 0
    
'     ' Screen updating turned off for faster execution
'     Application.ScreenUpdating = False
    
'     ' Loop through the Inventory rows (assuming headers are in Row 1, data starts at Row 2)
'     For i = 2 To lastRowInv
        
'         ' Read row values (Adjust column letters to match your actual Inventory layout)
'         itemId = wsInventory.Cells(i, "A").Value      ' Column A: Item ID
'         itemName = wsInventory.Cells(i, "B").Value    ' Column B: Item Name
'         currentGroup = wsInventory.Cells(i, "C").Value ' Column C: Product Group
'         currentBrand = wsInventory.Cells(i, "D").Value ' Column D: Product Brand
'         currentLoc = wsInventory.Cells(i, "E").Value   ' Column E: Location
'         currentWh = wsInventory.Cells(i, "F").Value    ' Column F: Warehouse
'         currentBranch = wsInventory.Cells(i, "G").Value ' Column G: Branch
'         itemPrice = wsInventory.Cells(i, "H").Value   ' Column H: Price/Amount
        
'         ' Check if the current item matches ALL specified criteria
'         If (currentGroup = ProdGroup) And _
'            (currentBrand = ProdBrand) And _
'            (currentLoc = Location) And _
'            (currentWh = Warehouse) And _
'            (currentBranch = Branch) Then
           
'             ' Check if adding this item stays within or completes our Target Amount
'             ' (Alternatively, remove "RunningTotal + itemPrice <= TargetAmount" if you want to get as close as possible even if it goes over slightly)
'             If RunningTotal + itemPrice <= TargetAmount Then
                
'                 ' Copy item details to Invoice Sheet (Adjust destination columns as needed)
'                 wsInvoice.Cells(nextRowInvc, "A").Value = itemId
'                 wsInvoice.Cells(nextRowInvc, "B").Value = itemName
'                 wsInvoice.Cells(nextRowInvc, "C").Value = itemPrice
                
'                 ' Update tracking metrics
'                 RunningTotal = RunningTotal + itemPrice
'                 nextRowInvc = nextRowInvc + 1
'                 MatchCount = MatchCount + 1
                
'                 ' Stop looping if we have perfectly hit or exhausted the budget target
'                 If RunningTotal >= TargetAmount Then Exit For
'             End If
            
'         End If
'     Next i
    
'     Application.ScreenUpdating = True
    
'     ' Output results or warnings to the user
'     If MatchCount = 0 Then
'         MsgBox "No items found matching the given criteria.", vbExclamation, "No Selection Made"
'     ElseIf RunningTotal < TargetAmount Then
'         MsgBox "Invoice created, but stock was insufficient to reach the target amount." & vbCrLf & _
'                "Target: " & FormatCurrency(TargetAmount) & vbCrLf & _
'                "Allocated: " & FormatCurrency(RunningTotal), vbByVal, "Partial Fulfillment"
'     Else
'         MsgBox "Successfully created invoice for " & FormatCurrency(RunningTotal), vbInformation, "Invoice Complete"
'     End If

' End Sub

' Sub TestInvoiceSelection()
'     ' Call the routine with sample arguments
'     ' Format: TargetAmount, ProductGroup, ProductBrand, Location, Warehouse, Branch
'     ' Call CreateInvoiceByCriteria(1500.0, "Electronics", "Sony", "North", "WH-02", "Branch-A")
' End Sub

''' =================================================================
''' Main function to auto-select items for an invoice
''' =================================================================
' get auto select items when create invoice by passing argument total amount of invoice, multi product group,multi product brand,multi location,multi warehouse,branch
' Public Function AutoSelectInvoiceItems( _
'     ByVal TargetAmount As Double, _
'     ByVal ProductGroups As String, _
'     ByVal ProductBrands As String, _
'     ByVal Locations As String, _
'     ByVal Warehouses As String, _
'     ByVal Branches As String) As Variant
    
'     Dim ws As Worksheet
'     Dim LastRow As Long, i As Long, matchCount As Long
'     Dim itemPrice As Double, currentTotal As Double
'     Dim groupArr() As String, brandArr() As String
'     Dim locArr() As String, whArr() As String, branchArr() As String
    
'     ' Split comma-separated arguments into arrays for multi-select matching
'     groupArr = Split(Trim(ProductGroups), ",")
'     brandArr = Split(Trim(ProductBrands), ",")
'     locArr = Split(Trim(Locations), ",")
'     whArr = Split(Trim(Warehouses), ",")
'     branchArr = Split(Trim(Branches), ",")
    
'     ' Set reference to your Inventory worksheet
'     Set ws = ThisWorkbook.Sheets("Inventory")
'     LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
'     ' Dynamic array to hold selected Item IDs
'     Dim SelectedItems() As String
'     ReDim SelectedItems(0)
'     matchCount = 0
'     currentTotal = 0#
    
'     ' Loop through inventory rows (assuming row 1 has headers, data starts at row 2)
'     For i = 2 To LastRow
'         ' 1. Validate multiple criteria using helper function
'         If IsMatch(ws.Cells(i, 3).Value, groupArr) And _
'            IsMatch(ws.Cells(i, 4).Value, brandArr) And _
'            IsMatch(ws.Cells(i, 5).Value, locArr) And _
'            IsMatch(ws.Cells(i, 6).Value, whArr) And _
'            IsMatch(ws.Cells(i, 7).Value, branchArr) Then
           
'             itemPrice = CDbl(ws.Cells(i, 8).Value) ' Assuming Column 8/H is Price

'             ' 2. Greedy allocation: check if adding this item fits the budget
'             If (currentTotal + itemPrice) <= TargetAmount Then
'                 ReDim Preserve SelectedItems(matchCount)
'                 SelectedItems(matchCount) = ws.Cells(i, 1).Value ' Column 1/A is Item ID
'                 matchCount = matchCount + 1
'                 currentTotal = currentTotal + itemPrice
'             End If

'             ' Exit early if we exactly hit or get within a negligible margin of the target
'             If Abs(currentTotal - TargetAmount) < 0.01 Then Exit For
'         End If
'     Next i
    
'     ' Return the array of selected items (or an error string if none found)
'     If matchCount > 0 Then
'         AutoSelectInvoiceItems = SelectedItems
'     Else
'         AutoSelectInvoiceItems = "No items matched criteria or target amount."
'     End If
' End Function

' ''' =================================================================
' ''' Helper function to check if a value exists within the multi-select array
' ''' =================================================================
' Private Function IsMatch(ByVal valToTest As String, ByRef searchArr() As String) As Boolean
'     Dim element As Variant
'     IsMatch = False
    
'     ' If the criteria array is empty or contains "*", bypass filtering for this criteria
'     If UBound(searchArr) < LBound(searchArr) Then
'         IsMatch = True
'         Exit Function
'     End If
'     If Trim(searchArr(0)) = "*" Or Trim(searchArr(0)) = "" Then
'         IsMatch = True
'         Exit Function
'     End If
    
'     ' Check for exact match across multi-select values
'     For Each element In searchArr
'         If UCase(Trim(valToTest)) = UCase(Trim(CStr(element))) Then
'             IsMatch = True
'             Exit Function
'         End If
'     Next element
' End Function

' Sub TestInvoiceSelection()
'     Dim results As Variant
'     Dim i As Long
    
'     ' Example: Target $500, passing comma-separated values for multi-selections
'     results = AutoSelectInvoiceItems(500.0, "Electronics,Home", "Sony,Samsung", "North,East", "WH-01", "Branch-A")
    
'     If IsArray(results) Then
'         MsgBox "Selected " & UBound(results) + 1 & " items for the invoice!"
'         For i = LBound(results) To UBound(results)
'             Debug.Print "Selected Item ID: " & results(i)
'         Next i
'     Else
'         MsgBox results
'     End If
' End Sub

' AI code VBA excel sheet get auto select items when create invoice by passing argument total amount of invoice, multi product group,multi product brand,mu
' Function AutoSelectInvoiceItems(ByVal TargetAmount As Double, _
'                                 ByVal ProductGroups As String, _
'                                 ByVal ProductBrands As String, _
'                                 ByVal Locations As String, _
'                                 ByVal Warehouses As String, _
'                                 ByVal Branch As String) As Variant
    
'     Dim wsInv As Worksheet
'     Dim lastRow As Long, i As Long
'     Dim currentAmount As Double
'     Dim itemPrice As Double, itemQty As Double, lineTotal As Double
    
'     ' Output array to hold selected item details
'     ' Adjust size or columns based on what you need to return
'     Dim selectedItems() As Variant
'     Dim itemCount As Long
'     itemCount = 0
'     currentAmount = 0#
    
'     ' 1. Set your source inventory sheet
'     On Error Resume Next
'     Set wsInv = ThisWorkbook.Sheets("Inventory")
'     On Error GoTo 0
    
'     If wsInv Is Nothing Then
'         MsgBox "Error: 'Inventory' sheet not found!", vbCritical
'         Exit Function
'     End If
    
'     lastRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row
    
'     ' 2. Loop through inventory items (Assuming headers are in row 1, data starts row 2)
'     ' Adjust column letters (e.g., "A", "B", "C") to match your exact inventory sheet layout
'     For i = 2 To lastRow
        
'         ' Check if we have already reached or exceeded the required invoice amount
'         If currentAmount >= TargetAmount Then Exit For
        
'         ' Extract criteria values from the current row
'         Dim rowGroup As String: rowGroup = wsInv.Cells(i, "A").Value      ' Column A: Product Group
'         Dim rowBrand As String: rowBrand = wsInv.Cells(i, "B").Value      ' Column B: Brand
'         Dim rowLoc As String: rowLoc = wsInv.Cells(i, "C").Value          ' Column C: Location
'         Dim rowWh As String: rowWh = wsInv.Cells(i, "D").Value            ' Column D: Warehouse
'         Dim rowBranch As String: rowBranch = wsInv.Cells(i, "E").Value    ' Column E: Branch
        
'         ' 3. Apply Multi-Filter Checks (using InStr to handle comma-separated multi-selects)
'         If (ProductGroups = "" Or InStr(1, ProductGroups, rowGroup, vbTextCompare) > 0) And _
'            (ProductBrands = "" Or InStr(1, ProductBrands, rowBrand, vbTextCompare) > 0) And _
'            (Locations = "" Or InStr(1, Locations, rowLoc, vbTextCompare) > 0) And _
'            (Warehouses = "" Or InStr(1, Warehouses, rowWh, vbTextCompare) > 0) And _
'            (Branch = "" Or LCase(rowBranch) = LCase(Branch)) Then

'             ' Extract item financial/quantity data
'             Dim itemID As String: itemID = wsInv.Cells(i, "F").Value      ' Column F: Item ID/SKU
'             itemQty = wsInv.Cells(i, "G").Value                           ' Column G: Available Qty
'             itemPrice = wsInv.Cells(i, "H").Value                         ' Column H: Unit Price
            
'             If itemQty > 0 And itemPrice > 0 Then
'                 lineTotal = itemQty * itemPrice
                
'                 ' Check if adding the whole lot exceeds the target amount
'                 If (currentAmount + lineTotal) > TargetAmount Then
'                     ' Calculate exactly how many pieces are needed to hit the target
'                     Dim neededAmount As Double
'                     neededAmount = TargetAmount - currentAmount
                    
'                     Dim neededQty As Double
'                     neededQty = Application.WorksheetFunction.RoundUp(neededAmount / itemPrice, 0)
                    
'                     ' If available stock covers the needed partial quantity
'                     If neededQty <= itemQty Then
'                         itemQty = neededQty
'                         lineTotal = itemQty * itemPrice
'                     End If
'                 End If
                
'                 ' Update total invoice accumulator
'                 currentAmount = currentAmount + lineTotal
'                 itemCount = itemCount + 1
                
'                 ' Resize array and store the selected item data
'                 ReDim Preserve selectedItems(1 To 4, 1 To itemCount)
'                 selectedItems(1, itemCount) = itemID      ' SKU
'                 selectedItems(2, itemCount) = itemQty     ' Quantity to pull
'                 selectedItems(3, itemCount) = itemPrice   ' Price
'                 selectedItems(4, itemCount) = lineTotal   ' Total for line
'             End If
            
'         End If
'     Next i
    
'     ' Return the populated array back to the calling sub
'     If itemCount > 0 Then
'         AutoSelectInvoiceItems = selectedItems
'     Else
'         AutoSelectInvoiceItems = Empty
'     End If
    
' End Function

' Sub CreateInvoice()
'     Dim invoiceData As Variant
'     Dim wsInvoice As Worksheet
'     Dim targetAmt As Double
'     Dim i As Long
    
'     Set wsInvoice = ThisWorkbook.Sheets("Invoice")
'     targetAmt = 5000.00 ' Your target invoice total amount
    
'     ' Pass comma-separated strings for multi-select arguments
'     invoiceData = AutoSelectInvoiceItems(targetAmt, _
'                                          "Electronics,Appliances", _
'                                          "Sony,Samsung", _
'                                          "North,East", _
'                                          "WH-01,WH-02", _
'                                          "Main Branch")
                                         
'     ' Check if items were found
'     If IsEmpty(invoiceData) Then
'         MsgBox "No items matched the criteria or stock is empty.", vbExclamation
'         Exit Sub
'     End If
    
'     ' Clear old invoice lines (assuming rows 5 onwards are item rows)
'     wsInvoice.Rows("5:100").ClearContents
    
'     ' Write the array data down onto the invoice template
'     For i = 1 To UBound(invoiceData, 2)
'         wsInvoice.Cells(4 + i, "A").Value = invoiceData(1, i) ' Item ID
'         wsInvoice.Cells(4 + i, "B").Value = invoiceData(2, i) ' Qty Allocated
'         wsInvoice.Cells(4 + i, "C").Value = invoiceData(3, i) ' Price
'         wsInvoice.Cells(4 + i, "D").Value = invoiceData(4, i) ' Total
'     Next i
    
'     MsgBox "Invoice populated successfully!", vbInformation
' End Sub

    
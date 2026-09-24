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
                If MyFSO.FolderExists(folder) = False Then
                    MyFSO.CreateFolder (folder)
                End If
            End If
            If (MyFSO.FolderExists(folder) = True) Then
                If (rng_file_setting.Cells(i,6) <> "") Then
                    subFolder1 = folder & "\" & rng_file_setting.Cells(i,6) 'report
                    If MyFSO.FolderExists(subFolder1) = False Then
                        MyFSO.CreateFolder (subFolder1)
                    End If
                End If
            End If
            If (MyFSO.FolderExists(subFolder1) = True) Then
                If (rng_file_setting.Cells(i,7) <> "") Then
                    subFolder2 = subFolder1 & "\" & rng_file_setting.Cells(i,7) 'subFolder2
                    If MyFSO.FolderExists(subFolder2) = False Then
                        ' debug.print i & " - folder name in report folder => " & MyFSO.FolderExists(subFolder2)
                        MyFSO.CreateFolder (subFolder2)
                    End If
                End If
            End If
            If (MyFSO.FolderExists(subFolder2) = True) Then
                If (rng_file_setting.Cells(i,8) <> "") Then
                    subFolder3 = subFolder2 & "\" & rng_file_setting.Cells(i,8) 'sub folder 3
                    If MyFSO.FolderExists(subFolder3) = False Then
                        ' debug.print i & " - sub folder 3 => " & MyFSO.FolderExists(subFolder3)
                        MyFSO.CreateFolder (subFolder3)
                    End If
                End If
            End If
            If (MyFSO.FolderExists(subFolder3) = True) Then
                If (rng_file_setting.Cells(i,9) <> "") Then
                    subFolder4 = subFolder3 & "\" & rng_file_setting.Cells(i,9) 'sub folder 4
                    If MyFSO.FolderExists(subFolder4) = False Then
                        ' debug.print i & " - sub folder 3 => " & MyFSO.FolderExists(subFolder4)
                        MyFSO.CreateFolder (subFolder4)
                    End If
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

    If (mainUomName = "") Then
        mainUomName = Application.VLOOKUP(mainUomId,rng_uom_list,5,FALSE)
    End If

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
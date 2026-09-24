Attribute VB_Name = "add_uom_conversion"

Public  Sub addUomConversion()
    OnStart

    Dim activeWorkbook As Workbook
    Dim ws_uom_con, ws_uom As Worksheet
    Dim rngUomCon,rngUomList  As Range
    Dim uomConName, mainUom, smallUom, middleUom,textMessageAlert As String
    Dim mainUomId, smallUomId, middleUomId, smallValUom, middleValUom, lastRowUomCon, uomConCount, isSmallUom, checkExistConversion,uomSelectedCount As Integer
    Dim arrUomValue, arrUomConId As Variant

    Set activeWorkbook  = Workbooks("index.xlsm")
    Set ws_uom_con      = activeWorkbook.Sheets("UomConversion")
    Set ws_uom          = activeWorkbook.Sheets("UoM")
    Set rngUomCon       = ws_uom_con.Range("C8:J8")
    Set rngUomList      = ws_uom.Range("C8:I" & ws_uom.Cells(Rows.Count,"C").End(XlUp).Row)

    isSmallUom       = 0
    uomConCount      = Application.Count(Range("G3:G4")) 'Application.WorksheetFunction.Count(Range("G3:G4"))
    uomSelectedCount = Application.COUNTA(Range("E3:F4"))

    mainUom      = MID(ActiveSheet.Range("uom_con_main_uom").Value,5,100)
    smallUom     = MID(ActiveSheet.Range("uom_con_small_uom").Value,5,100)
    middleUom    = MID(ActiveSheet.Range("uom_con_middle_uom").Value,5,100)

    mainUomId    = ActiveSheet.Range("uom_con_main_uom_id").Value
    smallUomId   = ActiveSheet.Range("uom_con_small_uom_id").Value
    middleUomId  = ActiveSheet.Range("uom_con_middle_uom_id").Value

    smallValUom  = ActiveSheet.Range("uom_con_small_value_uom").Value
    middleValUom = ActiveSheet.Range("uom_con_middle_value_uom").Value
    arrUomValue  = Split(ActiveSheet.Range("uom_con_array_value"),";")
    arrUomConId  = Split(ActiveSheet.Range("uom_con_array_con_id"),";")

    If (uomConCount > 1) Then
        If (smallValUom > 0 And middleValUom > 0) Then
            If (smallValUom Mod middleValUom) <> 0 Then 
                MsgBox "The middle Unit of Measure (UoM) must be divided by the small unit."
                exit sub
            End If
        End If
        If (middleValUom = "" Or middleValUom <=0) Then
            MsgBox "Middle uom value must be bigger than zero and divided by the small uom value"
            exit sub
        End If
    End If
    checkExistConversion = CheckExitUomConversion(CInt(mainUomId),1)
    k = 0
    If (checkExistConversion=0) Then
        For i = 1 To uomConCount
            lastRowUomCon = ws_uom_con.Cells(Rows.Count,"C").End(xlUp).Row - 6
            If CInt(arrUomConId(k)) > 0 And CInt(arrUomValue(k)) > 0 Then
                If (uomSelectedCount = 1) Then
                    isSmallUom = 1
                ElseIf (uomSelectedCount > 1) Then
                    If (i=uomConCount) Then 
                        isSmallUom = 1
                    Else
                        isSmallUom = 0
                    End If
                End If
                uomConName = Application.VLookup(CInt(arrUomConId(k)),rngUomList,5,False)
                rngUomCon.Cells(lastRowUomCon,1) = lastRowUomCon 'Id
                rngUomCon.Cells(lastRowUomCon,2) = mainUomId 'From Uom ID
                rngUomCon.Cells(lastRowUomCon,3) = CInt(arrUomConId(k)) 'To Uom ID
                rngUomCon.Cells(lastRowUomCon,4) = mainUom 'Main Uom
                rngUomCon.Cells(lastRowUomCon,5) = uomConName 'Uom Conversion
                rngUomCon.Cells(lastRowUomCon,6) = CInt(arrUomValue(k)) 'Value
                rngUomCon.Cells(lastRowUomCon,7) = isSmallUom 'Is Small Uom
                rngUomCon.Cells(lastRowUomCon,8) = 1 'Is Active
                debug.Print "uom id =" & CInt(arrUomConId(k)) & ",lastRowUomCon=" & lastRowUomCon
                k = k + 1
                textMessageAlert = "UoM conversion have been saved successful."  
            End If
        Next i
    ElseIf (checkExistConversion=1) Then
        textMessageAlert = "UoM conversion exist, Please try to input other UoM."
    End If
    
    OnEnd

    MsgBox textMessageAlert
    Exit Sub
End Sub

Private  Function CheckExitUomConversion(fromUomId As Integer,isActive As Integer) As Integer
    Dim ws_uom_con As Worksheet
    Dim i,lastUomConRow As Integer
    Dim checkExistTmp As Integer
    Dim rngUomCon As Range

    Set ws_uom_con  = Sheets("UomConversion")
    Set rngUomCon   = Sheets("UomConversion").Range("C8:J8")
    lastUomConRow   = ws_uom_con.Cells(Rows.Count, "C").End(xlUp).Row - 6
    checkExistTmp = 0
    For i = 1 To lastUomConRow
        If (fromUomId = rngUomCon.Cells(i, 2) And isActive = 1) Then
            checkExistTmp = 1
            Exit For
        End If
    Next i
    CheckExitUomConversion = checkExistTmp
End Function
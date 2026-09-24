Attribute VB_Name = "save_pdf_purchase_a4" 

Public  Sub savePdfPurchaseBillA4()
    OnStart
    checkFolderExist 'check not exist => create new folder
    Dim activeWorkbook As Workbook
    Dim ws_file_setting,ws_print_a4 As Worksheet
    Dim rng_file_setting,rng_print_a4 As Range
    Dim lastRowFileSetting As Integer
    Dim MyFSO As New FileSystemObject
    Dim directoryInvoice,invoiceCode,pdfFileName As String
    Dim fileExist As Boolean
    Dim confirmBoxAlert As VbMsgBoxResult

    confirmBoxAlert      = MsgBoxW("Are you want to export purchase bill as PDF", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Export PDF")
    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_file_setting  = activeWorkbook.Sheets("FilesSetting")
    Set ws_print_a4      = activeWorkbook.Sheets("PrintPurchaseBill")
    lastRowFileSetting   = ws_file_setting.Cells(Rows.Count,"C").End(xlUp).row

    Set rng_file_setting = ws_file_setting.Range("C8:M" & lastRowFileSetting)
    Set rng_print_a4     = ws_print_a4.Range("Print_Area")

    directoryInvoice     = Application.Vlookup(8,rng_file_setting,10,FALSE)
    invoiceCode          = ws_print_a4.Range("print_pb_code").Value
    
    If confirmBoxAlert = vbYes And checkDatabaseFile = TRUE Then
        If Len(Application.Vlookup(8,rng_file_setting,10,FALSE)) > 0 Then
            directoryInvoice = Left(directoryInvoice, Len(directoryInvoice) - 1)
            If MyFSO.FolderExists(directoryInvoice) = True Then
                pdfFileName  = invoiceCode & "_" & Format(ws_print_a4.Range("print_pb_date").Value, "dd-mm-yyyy") & ".pdf"
                fileExist    = checkFileIsExist(CStr(pdfFileName), CStr(directoryInvoice))
                If fileExist = True Then 
                    MsgBox "File purchase bill (" & CStr(pdfFileName) & ") already saved."
                    Exit Sub
                Else
                    rng_print_a4.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=directoryInvoice & "\" & pdfFileName, _
                    Quality:=xlQualityStandard, _
                    IgnorePrintAreas:=True, _
                    OpenAfterPublish:=True
                End If
            End If
        End If
    End If
    OnEnd
End Sub

Sub openPurchaseBillFolder()
    checkFolderExist 'check not exist => create new folder
    Dim activeWorkbook As Workbook
    Dim ws_file_setting As Worksheet
    Dim rng_file_setting As Range
    Dim directoryInvoice As String

    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_file_setting  = activeWorkbook.Sheets("FilesSetting")
    Set rng_file_setting = ws_file_setting.Range("C8:M" & ws_file_setting.Cells(Rows.Count,"C").End(xlUp).row)
    directoryInvoice     = Application.Vlookup(8,rng_file_setting,10,FALSE)
    directoryInvoice     = Left(directoryInvoice, Len(directoryInvoice) - 1)
    ' debug.Print "purchase bill =" & directoryInvoice
    ' Check if the folder actually exists before trying to open it
    If Dir(directoryInvoice, vbDirectory) <> "" Then
        Shell "explorer.exe """ & directoryInvoice & """", vbNormalFocus
    Else
        MsgBox "The folder path does not exist.", vbCritical, "Error"
    End If
End Sub

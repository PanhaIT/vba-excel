Attribute VB_Name = "save_pdf_quote_a4"

Private  Sub savePdfQuoteA4()
    OnStart
    checkFolderExist 'check not exist => create new folder
    Dim activeWorkbook As Workbook
    Dim ws_file_setting,ws_print_a4 As Worksheet
    Dim rng_file_setting,rng_print_a4 As Range
    Dim lastRowFileSetting As Integer
    Dim MyFSO As New FileSystemObject
    Dim directoryQuote,quoteCode,pdfFileName As String
    Dim fileExist As Boolean
    Dim confirmBoxAlert As VbMsgBoxResult

    confirmBoxAlert      = MsgBoxW("Are you want to export quotation as PDF", vbYesNoCancel  +  vbQuestion + vbDefaultButton1, "Export PDF")
    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_file_setting  = activeWorkbook.Sheets("FilesSetting")
    Set ws_print_a4      = activeWorkbook.Sheets("PrintQuotation")
    lastRowFileSetting   = ws_file_setting.Cells(Rows.Count,"C").End(xlUp).row

    Set rng_file_setting = ws_file_setting.Range("C8:M" & lastRowFileSetting)
    Set rng_print_a4     = ws_print_a4.Range("Print_Area")

    directoryQuote       = Application.Vlookup(13,rng_file_setting,10,FALSE)
    quoteCode            = ws_print_a4.Range("print_quote_code").Value
    If confirmBoxAlert = vbYes And checkDatabaseFile = TRUE Then
        If Len(Application.Vlookup(13,rng_file_setting,10,FALSE)) > 0 Then
            directoryQuote = Left(directoryQuote, Len(directoryQuote) - 1)
            If MyFSO.FolderExists(directoryQuote) = True Then
                pdfFileName  = quoteCode & "_" & Format(ws_print_a4.Range("print_quote_date").Value, "dd-mm-yyyy") & ".pdf"
                fileExist    = checkFileIsExist(CStr(pdfFileName), CStr(directoryQuote))
                If fileExist = True Then 
                    MsgBox "File invoice (" & CStr(pdfFileName) & ") already saved."
                    Exit Sub
                Else
                    rng_print_a4.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=directoryQuote & "\" & pdfFileName, _
                    Quality:=xlQualityStandard, _
                    IgnorePrintAreas:=True, _
                    OpenAfterPublish:=True
                End If
            End If
        End If
    End If
    OnEnd
End Sub

Private  Sub openQuoteFolder()
    checkFolderExist 'check not exist => create new folder
    Dim activeWorkbook As Workbook
    Dim ws_file_setting As Worksheet
    Dim rng_file_setting As Range
    Dim directoryQuote As String

    Set activeWorkbook   = Workbooks("index.xlsm")
    Set ws_file_setting  = activeWorkbook.Sheets("FilesSetting")
    Set rng_file_setting = ws_file_setting.Range("C8:M" & ws_file_setting.Cells(Rows.Count,"C").End(xlUp).row)
    directoryQuote     = Application.Vlookup(13,rng_file_setting,10,FALSE)
    directoryQuote     = Left(directoryQuote, Len(directoryQuote) - 1)
    ' debug.Print "directoryQuote =" & directoryQuote
    ' Check if the folder actually exists before trying to open it
    If Dir(directoryQuote, vbDirectory) <> "" Then
        Shell "explorer.exe """ & directoryQuote & """", vbNormalFocus
    Else
        MsgBox "The folder path does not exist.", vbCritical, "Error"
    End If
End Sub

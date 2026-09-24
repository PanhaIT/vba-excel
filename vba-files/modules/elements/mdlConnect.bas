Attribute VB_Name = "mdlConnect"
Public rs As New ADODB.Recordset
Public cn As New ADODB.Connection
Public SQL As String

' Sub Connect()
'     Set cn = New ADODB.Connection
'     Set rs = New ADODB.Recordset

'     Dim activeWorkbook, invSalePbWorkbook As Workbook
'     Dim ws_file_setting As Worksheet
'     Dim invSalePbPath, invSalePbFile As String

'     Set activeWorkbook = Workbooks("index.xlsm")
'     Set ws_file_setting = activeWorkbook.Sheets("FilesSetting")
    
'     If (checkDatabaseFile = True) Then
'         invSalePbPath = ws_file_setting.Range("database_path").Value
'         invSalePbFile = invSalePbPath & "kscm.accdb"
'     End If

'     With cn
'         .Provider = "Microsoft.ACE.OLEDB.12.0"
'         ' .ConnectionString = "Data Source = " & ThisWorkbook.Path & "\config\kscm.accdb" 'db.accdb
'         ' .ConnectionString = "Data Source =D:\OneDrive\YBA ADVISORY\B-Clients and Services\C024-KOK SNUOL\ks_system\database\kscm.accdb" 'db.accdb
'         ' D:\1_Documents\1_Companies\ks_system\database\
'         .ConnectionString = "Data Source = " & invSalePbFile  'db.accdb
'         .Open cn
'     End With
' End Sub

' Sub Disconnect()
'     Set rs = Nothing
'     Set cn = Nothing
' End Sub

' Public Sub OpenSystem()
'     frmMain.Show
' End Sub
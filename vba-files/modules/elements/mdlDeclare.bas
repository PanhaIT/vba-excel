Attribute VB_Name = "mdlDeclare"


' Public Nav                  As New clsNavMenu
' Public NavTopBar            As New clsTopBar
Public cBtn                 As New clsBtn
Public tbox                 As New clsTextBox
'Public cb                  As New clsCheckBox
' Public dash                 As New clsDashboard
' Public product              As New clsPerson
' Public orders               As New clsOrders
' Public customer             As New clsCustomer
' Public tb                   As New clsTabMenu
Public paging               As New clsPaging
'Public ob                  As New clsOptionButton

Public MenuCount            As Integer
Public mFrameWidth          As Integer

'Public PassiveTabLabel      As MSForms.Label
'Public ActiveTabLabel       As MSForms.Label
Public SortDirection        As String

Sub GarbageCollector()
    Set Nav = Nothing
    ' Set NavTopBar = Nothing
    Set cBtn = Nothing
    Set tbox = Nothing
    Set cb = Nothing
    ' Set dash = Nothing
    ' Set product = Nothing
    ' Set orders = Nothing
    Set tb = Nothing
    Set paging = Nothing
    Set puantaj = Nothing
    Set ob = Nothing

    ' Set customer = Nothing

    SortDirection = "asc"
    FormLoad = False
    Disconnect
End Sub


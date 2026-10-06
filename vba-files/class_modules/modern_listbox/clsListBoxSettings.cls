VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListBoxSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' In theory this should be treated as a singleton class. That means only on instance is created.
' However making it an instance means creating private(i.e. global variable) which has to be explicitly deleted.
' Should be done in terminate of the clsModernListbox class. However terminate is not fired when the class in
' publicNotCreatable.
' Also if you cancel the QueryClose on the userform ,which is standard, then UserForm terminate is not called unless
' you explicitly call it.
' The ultimate issue is that the classes are not dying and sometimes it leads to an catastrophic errors as the code
' tries to access a frame that no longer exists.
 

Public Enum sortOrder
    soNeutral = 0
    soAscending = 1
    soDescending = -1
End Enum

Public checkboxPadding As Long
Public checkboxLeft As Long
Public dataFontColor As Long
Public dataFontSize As Long
Public dataFontName As String
Public dataRowPaddingTop As Long
Public dataRowPaddingBottom As Long
Public recordsPerPage As Long
Public columnHeaderBackgroundColor As Long
Public columnHeaderFontColor As Long
Public columnHeaderFontSize As Long
Public columnHeaderFontName As String
Public headerRowPaddingTop As Long
Public headerRowPaddingBottom As Long
Public hoverColorheader As Long
Public hoverColorRow As Long
Public sortSymbolFontSize As Long
Public sortSymbolFontName As String
Public labelWidthDefault As Long
Public labelNoRecordWidth As Long
Public labelPadding As Long
Public listboxBackgroundColor As Long
Public frameTopPadding As Long
Public frameLeftPadding As Long
Public frameRightPadding As Long
Public frameBottomPadding As Long
Public frameHeightMinimum As Long
Public frameWidthMinimum As Long
Public paginationLabelFontName As String
Public paginationLabelFontSize As Long
Public paginationLabelPadding As Long
Public paginationLabelHeight As Long
Public paginationLabelWidth As Long
Public paginationPaddingRight As Long
Public paginationButtonLimit
Public rowFramePaddingLeft As Long
Public rowFramePaddingRight As Long
Public rowSelectionColor As Long
Public searchBoxFontName As String
Public searchBoxFontSize As Long
Public searchBoxHeight As Long
Public searchBoxWidth As Long
Public searchBoxPaddingBottom As Long
Public searchBoxPaddingRight As Long
Public searchBoxLabelWidth As Long
Public searchBoxLabelFontSize As Long
Public searchBoxLabelFontName As String
Public searchTurnedOn As Boolean
Public scrollHeightPadding As Long

Private Sub Class_Initialize()
    checkboxPadding = 0
    checkboxLeft = 0
    dataFontColor =  rgb(32, 32, 32)
    dataFontName = "Khmer OS Battambang"
    dataFontSize = 11
    dataRowPaddingTop = 6
    dataRowPaddingBottom = 10
    recordsPerPage = 15
    columnHeaderBackgroundColor = rgb(221, 221, 221)
    columnHeaderFontColor =  rgb(36, 36, 36)
    columnHeaderFontName = "Khmer OS Battambang"
    columnHeaderFontSize = 11
    headerRowPaddingTop = 6
    headerRowPaddingBottom = 12
    hoverColorheader =  rgb(236, 238, 240)
    hoverColorRow = rgbLightGray
    labelWidthDefault = 100
    labelNoRecordWidth = 200
    labelPadding = 0
    listboxBackgroundColor =  rgb(255, 255, 255)
    frameTopPadding = 5
    frameLeftPadding = 0
    frameRightPadding = 0
    frameBottomPadding = 5
    frameHeightMinimum = 100
    frameWidthMinimum = 100
    paginationLabelFontName = "Khmer OS Battambang"
    paginationLabelFontSize = 11
    paginationLabelPadding = 15
    paginationLabelHeight = paginationLabelFontSize * 1.5
    paginationLabelWidth = 10
    paginationPaddingRight = 10
    paginationButtonLimit = 10
    rowSelectionColor =  rgb(240, 240, 240)
    searchBoxFontName = "Khmer OS Battambang"
    searchBoxFontSize = 11
    searchBoxHeight = 30
    searchBoxWidth = 150
    searchBoxPaddingBottom = 1
    searchBoxPaddingRight = 0
    searchBoxLabelWidth = 50
    searchBoxLabelFontSize = 11
    searchBoxLabelFontName = "Khmer OS Battambang"
    searchTurnedOn = True
    sortSymbolFontSize = 11
    sortSymbolFontName = "myicons" 'Wingdings 3
    scrollHeightPadding = 0
    rowFramePaddingLeft = 0
    rowFramePaddingRight = 5
End Sub

Public Property Get GetCheckboxLeft() As Long
    GetCheckboxLeft = checkboxLeft
End Property
Public Property Get getCheckboxHeight() As Long
    getCheckboxHeight = (getDataLabelHeight() / 1.5) + 4
End Property
Public Property Get getCheckboxTop() As Long
    getCheckboxTop = getDataLabelTop() + checkboxPadding
End Property
Public Property Get getCheckboxWidth() As Long
    getCheckboxWidth = getCheckboxHeight()
End Property

Public Property Get GetSelectWidth(selectType As fmMultiSelect) As Long
    GetSelectWidth = IIf(selectType = fmMultiSelectSingle, 0, getCheckboxWidth() * 2)
End Property

Public Property Get getRowBackgroundColor() As Long
    getRowBackgroundColor = rgbWhite
End Property

Public Property Get getDataFontUnderline() As Boolean
    getDataFontUnderline = False
End Property
Public Property Get getDataLabelHeight() As Long
    getDataLabelHeight = dataFontSize * 1.5  + 4 'labelHeight
End Property
Public Property Get getDataLabelPaddingLeft() As Long
    getDataLabelPaddingLeft = labelPadding
End Property
Public Property Get getDataLabelTextAlign() As fmTextAlign
    getDataLabelTextAlign = fmTextAlignLeft
End Property
Public Property Get getDataLabelTop() As Long
    getDataLabelTop = dataRowPaddingTop
End Property

Public Property Get getRecordsPerPage() As Long
    getRecordsPerPage = recordsPerPage
End Property

'Public Property Get getParentFrameBackgroundColor() As Long
'    getParentFrameBackgroundColor = rgbWhite
'End Property

Public Property Get getColumnHeaderLabelHeight() As Long
    getColumnHeaderLabelHeight = columnHeaderFontSize * 1.5
End Property
Public Property Get getColumnHeaderRowHeight() As Long
    getColumnHeaderRowHeight = getColumnHeaderLabelHeight() + (headerRowPaddingTop + headerRowPaddingBottom)
End Property
Public Property Get getColumnHeaderLabelTop() As Long
    getColumnHeaderLabelTop = headerRowPaddingTop
End Property
'Public Property Get getHighlightColorSelection() As Long
'    getHighlightColorSelection = rgbLightBlue
'End Property

Public Property Get GetLabelSortSymbol(ByVal sortDirection As sortOrder) As String
    Select Case sortDirection
        Case soNeutral: GetLabelSortSymbol = ChrW(&H0055)
        Case soAscending: GetLabelSortSymbol = ChrW(&HE091)
        Case soDescending: GetLabelSortSymbol = ChrW(&HE092)
    End Select
End Property
Public Property Get GetLabelSortSymbolWidth() As Long
    GetLabelSortSymbolWidth = 12
End Property
Public Property Get GetLabelSortSymbolPaddingRight() As Long
    GetLabelSortSymbolPaddingRight = 6
End Property
Public Property Get getLabelSortSymbolFontName() As String
    getLabelSortSymbolFontName = sortSymbolFontName
End Property
Public Property Get getLabelSortSymbolFontSize() As Long
    getLabelSortSymbolFontSize = sortSymbolFontSize
End Property

Public Property Get getLabelNoRecordWidth() As Long
    getLabelNoRecordWidth = labelNoRecordWidth
End Property

Public Property Get getframeTopPadding() As Long
    getframeTopPadding = frameTopPadding
End Property
Public Property Get getframeBottomPadding() As Long
    getframeBottomPadding = frameBottomPadding
End Property

Public Property Get getFrameLeftPadding() As Long
    getFrameLeftPadding = frameLeftPadding
End Property
Public Property Get getFrameRightPadding() As Long
    getFrameRightPadding = frameRightPadding
End Property
Public Property Get getFrameHeightMinimum() As Long
    getFrameHeightMinimum = frameHeightMinimum
End Property
Public Property Get getFrameWidthMinimum() As Long
    getFrameWidthMinimum = getsearchBoxLabelWidth() + getsearchBoxLabelWidth() + getSearchBoxPaddingRight()
End Property
Public Property Get getDataRowHeight() As Long
    getDataRowHeight = getDataLabelHeight() + (dataRowPaddingTop + dataRowPaddingBottom)
End Property
Public Property Get getLabelWidthDefault() As Long
    getLabelWidthDefault = labelWidthDefault
End Property
Public Property Get getRowFramePaddingLeft() As Long
    getRowFramePaddingLeft = rowFramePaddingLeft
End Property
Public Property Get getRowFramePaddingRight() As Long
    getRowFramePaddingRight = rowFramePaddingRight
End Property
Public Property Get getPaginationLabelFontColor() As Long
    getPaginationLabelFontColor = vbBlue
End Property
Public Property Get getPaginationLabelFontColorGrayed() As Long
    getPaginationLabelFontColorGrayed = rgbGrey
End Property
Public Property Get getPaginationLabelFontName() As String
    getPaginationLabelFontName = paginationLabelFontSize
End Property
Public Property Get getPaginationLabelFontSize() As Long
    getPaginationLabelFontSize = paginationLabelFontSize
End Property
Public Property Get getPaginationLabelFontUnderline() As Boolean
    getPaginationLabelFontUnderline = True ' xlUnderlineStyleSingle
End Property
Public Property Get getPaginationLabelPadding() As Long
    getPaginationLabelPadding = paginationLabelPadding
End Property
Public Property Get getPaginationLabelHeight() As Long
    getPaginationLabelHeight = paginationLabelHeight
End Property
Public Property Get getPaginationLabelTextAlign() As fmTextAlign
    getPaginationLabelTextAlign = fmTextAlignCenter
End Property
Public Property Get getPaginationLabelWidth() As Long
    getPaginationLabelWidth = paginationLabelWidth
End Property
Public Property Get getPaginationPaddingRight() As Long
    getPaginationPaddingRight = paginationPaddingRight
End Property
Public Property Get getPaginationButtonLimit() As Long
    getPaginationButtonLimit = paginationButtonLimit
End Property
Public Property Get getScrollHeightPadding() As Long
    getScrollHeightPadding = scrollHeightPadding
End Property
Public Property Get getSearchBoxHeight() As Long
    getSearchBoxHeight = searchBoxHeight
End Property
Public Property Get getSearchBoxWidth() As Long
    getSearchBoxWidth = searchBoxWidth
End Property
Public Property Get getSearchBoxPaddingRight() As Long
    getSearchBoxPaddingRight = searchBoxPaddingRight
End Property
Public Property Get getSearchBoxPaddingBottom() As Long
    getSearchBoxPaddingBottom = searchBoxPaddingBottom
End Property
Public Property Get getsearchBoxLabelWidth() As Long
    getsearchBoxLabelWidth = searchBoxLabelWidth
End Property
Public Property Get getSearchTurnedOn() As Long
    getSearchTurnedOn = searchTurnedOn
End Property








VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModernListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ExcelMacroMastery.com
' Version 2.0 Updated 15-05-2025
' Guide to the Modern Listbox: https://excelmacromastery.com/modern-vba-listbox/
'
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: The main class of the Listbox. Everything that happens is managed through here

' Set Column Widths
' 1. If column settings not supplied then use SetColumnDimensionsDefaults
' 2A. If column settings supplied then use SetColumnDimensions
' 2B.       If settings not supplied for all columns use default for rest

Private Const NOT_SET As Long = -1

Public Event ItemSelected(row As Long)
Public Event ItemDeselected(row As Long)
Public Event ItemDoubleClicked(row As Long)

Enum listboxErrors
    errObjectNotSet = vbObjectError + 513
    errParameterNotArray
    errColumnWidthFormatInvalid
End Enum

' Events
Private m_checkboxEvents() As clsEventsRowCheckbox
Private m_checkboxHeaderEvents As New clsEventsHeaderCheckbox
Private m_columnDimensions() As ColumnDimension ' columns left position and widths
Private m_frameMain As New clsEventsFrameMain
Private m_frameRowEvents() As clsEventsRowFrame
Private m_frameRowHeaderEvents As New clsEventsRowHeaderFrame
Private m_headerLabelEvents() As clsEventsHeaderLabel
Private m_labelRowEvents() As clsEventsRowLabel ' used for creating the mouseover event for the rows
Private m_pageNumberEvents() As clsEventsPaginationNumbers
Private m_pageArrowEvents() As clsEventsPaginationArrows
Private m_searchTextboxEvents(1 To 1) As New clsEventsSearchTextbox
Private m_searchLabelEvents(1 To 1) As New clsEventsSearchLabel
Private m_sortEvents() As clsEventsHeaderSymbol ' used for creating the sort events for the headers
Private m_visible As Boolean ' stores the visibility state of the control

' Helper classes
Private m_controlAccessors As New clsControlAccessors
Private m_controlAttributes As clsControlAttributes
Private m_nameHelper As New clsControlNameHelpers
Private m_calcs As clsControlPositionCalcs
Private m_settings As New clsListBoxSettings
Private m_hoverRow As clsHoverRow
Private m_hoverHeader As clsHoverHeader
Private m_highlightRow As clsHighlightRow
Private m_highlightTracker As New clsHighlightTracker
Private m_sortTracker As New clsSortTracker

' Helper variables
Private m_data As Variant ' The data in the listbox not including the header row
Private m_dataStore As Variant ' use as backup for m_data when filtering
Private m_firstBuild As Boolean ' we create a new parentframe if not first build(to avoid automation error)
Private m_filterString As String

Private m_pageNumber As Long
Private m_pageCount As Long
Private m_originalParentFrame As MSForms.frame
Private m_parentFrameManager As New clsParentFrameManager

Private m_searchOn As Boolean
Private m_AutomaticHeight As Boolean
Private m_AutomaticWidth As Boolean

' Set externally using properties
Private m_headerFields As Variant
Private m_hoverOn As Boolean
Private m_ID As Long ' 07-05-2025 Unique ID for each Listbox created so each searchbox has a unique name
Private m_multiSelect As fmMultiSelect
Private m_recordsPerPage As Long

' EVENTS
Private Sub Class_Initialize()
    m_searchOn = m_settings.getSearchTurnedOn()
    m_recordsPerPage = m_settings.getRecordsPerPage()
    m_firstBuild = True
    m_visible = True
End Sub

' PROPERTIES
Public Property Get ColumnCount() As Long
    ColumnCount = GetDataColumnCount()
End Property

Public Property Let columnWidths(ByVal columnWidths As String)
    If IsColumnWidthStringValid(columnWidths) = False Then
        Err.Raise errColumnWidthFormatInvalid _
                   , "clsModernListbox.columnWidths property" _
                   , "Column widths are invalid. Please enter a valid string and try again."
        Exit Property
    End If
    Call SetColumnDimensions(columnWidths)
    If IsListboxBuilt() = False Then Exit Property
    
    Call ResizeColumns
    Call ResizeHeaders
    Call SetScrollBars
    Call ResizeSearch
    Call ResizePagination
End Property

' 07-05-2025 Returns all the controls e.g. frame, searchbox, pagination buttons etc.
'Public Property Get Controls() As Collection
'
'    Dim ListboxControls As New Collection
'    ListboxControls.Add m_parentFrame
'    ListboxControls.Add m_controlAccessors.GetSearchBox(m_ID)
'    ListboxControls.Add m_controlAccessors.GetSearchLabel(m_ID)
'
'    Dim i As Long
'    For i = 1 To UBound(m_pageNumberEvents, 1)
'        ListboxControls.Add m_controlAccessors.GetPaginationNumberLabel(m_ID, i)
'    Next i
'
'    For i = 1 To UBound(m_pageArrowEvents, 1)
'        ListboxControls.Add m_controlAccessors.GetPaginationArrowLabel(m_ID, i)
'    Next i
'    Set Controls = ListboxControls
'
'End Property

Public Property Let HeaderFieldsFromArray(ByRef myHeaderData As Variant)
    Call SetHeaderFieldsFromArray(myHeaderData)
End Property
Public Property Let HeaderFieldsFromString(myHeaderData As String)
    Call SetHeaderFieldsFromArray(StringTo2DArray(myHeaderData, ";"))
End Property

Public Property Get Height() As Long
    Height = m_parentFrameManager.parentFrame.Height
End Property
Public Property Let Height(newHeight As Long)
    If newHeight < m_settings.getFrameHeightMinimum() Then Exit Property
    m_parentFrameManager.parentFrame.Height = newHeight
    Call RebuildPagination
End Property

Public Property Let ID(argID As Long)
    m_ID = argID
End Property
Public Property Get width() As Long
    width = m_parentFrameManager.parentFrame.width
End Property
Public Property Let width(newWidth As Long)
    If newWidth < m_settings.getFrameWidthMinimum() Then Exit Property
    m_parentFrameManager.parentFrame.width = newWidth
    Call RebuildPagination
    Call RebuildSearch
End Property

Public Property Let HoverOn(newHoverOn As Boolean)
    m_hoverOn = newHoverOn
    If newHoverOn = False Then
        Call ClearHover
    End If
End Property

Public Property Let List(ByRef data As Variant)
    Call AddDataToListBox(data)
End Property

Public Property Let multiSelect(selectType As fmMultiSelect)
    m_multiSelect = selectType
    If IsListboxBuilt() = False Then Exit Property
    Call SetMultiSelectControls
    If m_multiSelect = fmMultiSelectSingle Then
        Call ClearSelections
    End If
End Property

Public Property Get multiSelect() As fmMultiSelect
    multiSelect = m_multiSelect
End Property

'Public Property Get parentFrame() As MSForms.frame
'    Set parentFrame = m_parentFrameManager.parentFrame
'End Property

Public Property Set parentFrame(ByVal myParentFrame As MSForms.frame)
    
    
    myParentFrame.Visible = False
    Set m_parentFrameManager.ParentFrameOriginal = myParentFrame
    
    Call m_parentFrameManager.InitializeFrame(m_settings, Me)
    Call AssignFrame
    Call InitializePageNumber

End Property

Public Property Let recordsPerPage(ByRef newRecordsPerPage As Long)
    Call ChangeRecordsPerPage(newRecordsPerPage)
    Call ResizeHeight
    Call RebuildPagination
    Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)
End Property

Public Property Let ScrollBarSize(newScrollWidth As Long)
    m_parentFrameManager.parentFrame.ScrollWidth = newScrollWidth
End Property

Public Property Get ScrollBars() As fmScrollBars
    ScrollBars = m_parentFrameManager.parentFrame.ScrollBars
End Property
Public Property Let ScrollBars(newScrollBars As fmScrollBars)
    m_parentFrameManager.parentFrame.ScrollBars = newScrollBars
    Call SetScrollBars
End Property

Public Property Get SelectedItem() As Variant
    Call getSelectedItem(SelectedItem)
End Property

Public Property Get SelectedItems() As Variant
    Call getSelectedItems(SelectedItems)
End Property

Public Property Get SelectedItemsCount() As Long
    SelectedItemsCount = m_highlightTracker.SelectedRowCount()
End Property

Public Property Let AutomaticHeight(turnOn As Boolean)
    m_AutomaticHeight = turnOn
    If turnOn = False Then Exit Property
    Call ResizeHeight
    Call RebuildPagination
End Property
Public Property Get AutomaticHeight() As Boolean
    AutomaticHeight = m_AutomaticHeight
End Property

Public Property Let AutomaticWidth(turnOn As Boolean)
    m_AutomaticWidth = turnOn
    If turnOn = False Then Exit Property
    Call ResizeWidth
    Call RebuildPagination
    Call ResizeSearch
End Property

Public Property Get AutomaticWidth() As Boolean
    AutomaticWidth = m_AutomaticWidth
End Property

' 07-05-2025
Public Property Get Visible() As Boolean
    Visible = m_visible
End Property

Public Sub SetVisible(visibleState As Boolean, Optional AllControls As Boolean = False)
    m_parentFrameManager.parentFrame.Visible = visibleState
    If AllControls = True Then
        m_controlAccessors.GetSearchBox(m_ID).Visible = visibleState
        m_controlAccessors.GetSearchLabel(m_ID).Visible = visibleState
        Call SetPaginationNumberLabelsVisible(visibleState)
        Call SetPaginationArrowLabelsVisible(visibleState)
    End If
    m_visible = visibleState
End Sub

' END OF PROPERTIES
Public Sub ClearSelections()
    If IsListboxBuilt() = False Then Exit Sub
    Call m_highlightTracker.ClearAll
    Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)
    Call m_highlightRow.SetCheckHeader(False)
End Sub

Public Function GetRow(index As Long) As Variant
    Dim dataRow As Variant
    Call getRowData(dataRow, index)
    GetRow = dataRow
End Function

Public Sub SelectAll()
    Call m_highlightTracker.SelectAll(m_data, GetDataRowCount)
    Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)
    Call m_highlightRow.SetCheckHeader(True)
End Sub

' Turn selection for a row on or off
Public Sub SetSelected(index As Long, turnedOn As Boolean)
    Call SetSelectRow(index, turnedOn)
End Sub

' Clear all data from the listbox
Public Sub Clear()
    Call RemoveHeaderControls
    Call RemoveAllDataControls
    Call RemovePaginationButtons
    Call RemovePaginationArrows
    Call EnableSearchBox(False)
End Sub

' END OF METHODS

Private Sub SetHeaderFieldsFromArray(myHeaderData As Variant)
    'm_headerFields = myHeaderData
    If IsArray(myHeaderData) = False Then
            Err.Raise errParameterNotArray, "modernListbox.SetHeaderFieldsFromArray" _
                , "The parameter myHeaderData must be assigned to an array. " _
                & vbNewLine & "If you are using a range then add .Value. " _
                & vbNewLine & "If you are using a string then use the Split function to convert to an array."
    End If
    Call SetHeaderFieldsFromCaller(myHeaderData)
    If IsListboxBuilt() = False Then Exit Sub
    Call FillHeaderControlsWithData

End Sub

Private Sub AddDataToListBox(ByRef data As Variant)
    If IsArrayAllocated(data) = False Then
            Err.Raise errParameterNotArray, "Property Let List" _
        , "The List property of the ListBox must be assigned to an array. If you are using a range then add .Value"

    End If

    Call CopyData(data)
    Call BackupData
    Call BuildListBox

End Sub

Private Sub CopyData(data As Variant)
    
    ' Add ID row so we can highlight after sorting
    ReDim m_data(LBound(data, 1) To UBound(data, 1), LBound(data, 2) To UBound(data, 2) + 1)
    
    ' Copy array
    Dim i As Long, j As Long
    Dim column As Long
    For i = LBound(data, 1) To UBound(data, 1)
        m_data(i, 1) = i ' Store ID
        column = 1
        For j = GetFirstDataColumn() To GetLastDataColumn()
            m_data(i, j) = data(i, column)
            column = column + 1
        Next j
    Next i
    
End Sub

' Returns all selected items as a variant aray
Private Sub getSelectedItems(items As Variant)
    If m_highlightTracker.SelectedVisibleRowCount() = 0 Then Exit Sub
    Dim outSelectedRows As Variant
    Call arrayCopyNonSequentialRows(m_data, m_highlightTracker.SelectedRowNumbers(), outSelectedRows)
    items = outSelectedRows
End Sub

' Returns currently selected item
Private Sub getSelectedItem(item As Variant)
    If m_highlightTracker.SelectedVisibleRowCount() = 0 Then Exit Sub

    ReDim item(1 To 1, 1 To GetDataColumnCount())
    Call arrayCopyRow(item, 1, m_data, m_highlightTracker.SelectedFirstVisibleRow, startColumn:=2)

End Sub

Private Sub getRowData(rowData As Variant, realRow As Long)
     ReDim rowData(1 To 1, 1 To GetDataColumnCount())
    Call arrayCopyRow(rowData, 1, m_data, realRow, startColumn:=2)
End Sub

Private Sub SetHeaderFields()
    Call SetDefaultHeaderFields
End Sub

Private Sub SetDefaultHeaderFields()

    If IsEmpty(m_headerFields) = False Or IsEmpty(m_data) = True Then Exit Sub
    
    ReDim m_headerFields(1 To 1, 1 To GetDataColumnCount())
    Dim j As Long
    For j = LBound(m_headerFields, 2) To UBound(m_headerFields, 2)
        m_headerFields(1, j) = "Column " & j
    Next j
End Sub

Private Sub SetHeaderFieldsFromCaller(myHeaderData As Variant)
    ' If headers are supplied for every column then the simply assign the array
    If UBound(myHeaderData, 2) >= GetDataColumnCount() Then
        m_headerFields = myHeaderData
    Else
        ' If less headers than columns then fill the ones supplied. The others will contain default or
        ' pre existing headers
        Dim j As Long
        For j = LBound(myHeaderData, 2) To UBound(myHeaderData, 2)
            m_headerFields(1, j) = myHeaderData(1, j)
        Next j
    End If
End Sub

Private Sub AssignFrame()
    
    Set m_controlAccessors.parentFrame = m_parentFrameManager
    
    Set m_calcs = CreateControlPositionCalcs(m_parentFrameManager, m_settings)
    Set m_controlAttributes = CreateControlAttributes(m_settings, m_calcs)
    Set m_highlightRow = CreateHighlightRow(m_parentFrameManager, m_settings)
    Set m_hoverHeader = CreateHoverHeader(m_parentFrameManager, m_settings)
    Set m_hoverRow = CreateHoverRow(m_parentFrameManager, m_settings)
End Sub

Private Sub InitializePageNumber()
    m_pageNumber = 1
End Sub

' MANAGER SUBS

' Add data controls, events and values
Private Sub BuildListBox()

    If m_firstBuild = False Then
        Call m_parentFrameManager.InitializeFrame(m_settings, Me)
    End If
    m_pageCount = GetPageCount
    
    Call SetHeaderFields ' first as may use header from data
    
    Call SetColumnDimensionDefaults

    ' Remove existing controls
    Call RemoveHeaderControls
    Call RemoveAllDataControls
    
    ' Header Controls
    Call AddHeaderControls
    Call SetColumnSortStatus
    Call AddMultiSelectControlsHeader
    Call ShowMultiSelectControlsHeader
    
    ' Add row controls
    Call AddRowControls
    Call HideUnusedRows
    Call AddMultiSelectControlsRow
    Call ShowMultiSelectControlsRow
    
    ' Add data to controls
    Call FillControlsWithData
    Call FillHeaderControlsWithData
    
    ' Create events
    Call CreateDataControlHoverEvents
    Call CreateHeaderControlHoverEvents

    
    ' Add search and pagination
    Call SetScrollBars
    
    ' Set height/width automatically if set
    AutomaticHeight = m_AutomaticHeight
    AutomaticWidth = m_AutomaticWidth
    
    Call AddSearchBox
    Call RebuildPagination
    
    m_firstBuild = False
End Sub

Private Sub ChangeRecordsPerPage(newRecordsPerPage)
    m_recordsPerPage = newRecordsPerPage
    m_pageCount = GetPageCount
    m_pageNumber = 1
    If IsListboxBuilt() = False Then Exit Sub
    
    
    Call RemoveAllDataControls
    Call AddRowControls
    Call AddMultiSelectControlsRow
    Call ShowMultiSelectControlsRow
    
    Call FillControlsWithData
    Call HideUnusedRows
    
    Call CreateDataControlHoverEvents
    
    Call SetScrollBars
End Sub

Private Sub RebuildPagination()
    If IsListboxBuilt() = False Then Exit Sub
    If paginationExists() = True Then Exit Sub
    Call RemovePaginationButtons
    Call RemovePaginationArrows
    Call AddPaginationButtons
    Call AddPaginationArrows
    Call ResizePagination
End Sub

Private Sub RebuildSearch()
    Call RemoveSearch
    Call AddSearchBox
End Sub


Private Sub SetMultiSelectControls()
    Call ResizeColumns
    Call ResizeHeaders
    Call SetScrollBars
    Call ResizeSearch
    Call ResizePagination
    Call ShowMultiSelectControlsHeader
    Call ShowMultiSelectControlsRow
    

End Sub

Private Sub AddSearchBox()
    If m_searchOn = False Then Exit Sub
    
    If searchBoxExists() = True Then
        ' Uses this when list has been cleared as searchbox is disabled and not deleted
        Call EnableSearchBox(True)
    Else
        Call CreateSearch
    End If
End Sub

Private Sub EnableSearchBox(enabled As Boolean)
    m_controlAccessors.GetSearchBox(m_ID).enabled = enabled
End Sub

Private Function searchBoxExists() As Boolean
    searchBoxExists = m_controlAccessors.CheckSearchBoxExists(m_ID)
End Function
Private Function paginationExists() As Boolean
    paginationExists = m_controlAccessors.CheckPaginationExists(m_ID)
End Function


Private Sub AddPaginationButtons()
    Call AddPaginationNumberLabels
    Call SetPaginationNumberCaptions
    Call SetPaginationNumberLabelStatus
    
End Sub

Private Sub RemovePaginationButtons()
    Call RemovePaginationNumberLabels
End Sub

Private Sub AddPaginationArrows()
    Call AddPaginationArrowLabels
    Call SetPaginationArrowLabels
End Sub

Private Sub RemovePaginationArrows()
    Call RemovePaginationArrowLabels
End Sub

' Friend makes it only available in this project. We don't want to make it available to UserForms that use
' the clsModernList object
Friend Sub FilterRows(ByVal text As String)
    
    If IsListboxBuilt() = False Then Exit Sub
    
    Call RemovePaginationButtons
    Call RemovePaginationArrows
    Call RecallData
    
    Dim filteredRowCount As Long
    filteredRowCount = FilterData(text)
    
    If filteredRowCount > 0 Then
        ' If records found then display filtered records
        m_pageNumber = 1
        m_pageCount = GetPageCount
        Call FillControlsWithData
        Call ShowFilteredControlsOnly
        Call ShowRowLabels(Visible:=True)
        Call m_highlightTracker.ResetRecordIDs(m_data)
        Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)
        Call AddPaginationButtons
        Call AddPaginationArrows
        Call SetPaginationNumberCaptions
        Call ResizePagination
    Else
        ' If no records found then hide all data except first label. Display message in first label.
        Call ShowRowLabels(Visible:=False)
        Call DisplayMessageInRow
        Call HideUnusedRowsNoRecords
    End If
    
End Sub

' Friend makes Sort available to the event class but not to userform that uses clsModernListbox
Friend Sub Sort(labelName As String)

    Dim column As Long: column = m_nameHelper.GetColumnFromHeaderLabelName(labelName)

    Dim dataFiltered As Boolean: dataFiltered = IsDataFiltered()
    If dataFiltered = True Then
        Call RecallData
    End If
    Call SortData(column)
    Call BackupData
    Call SetSortSymbol(column)

    If dataFiltered = True Then
        Call FilterData(m_filterString) ' re apply filter
    End If
    Call FillControlsWithData

    Call UpdateHighlightedRecords(column)

End Sub


Private Sub UpdateHighlightedRecords(column As Long)

    Call m_highlightTracker.ResetRecordIDs(m_data)
    Call m_highlightTracker.SortSelectedItems(m_sortTracker.GetSortDirection(column))
    Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)
    Call m_highlightRow.SetCheckHeader(False)
End Sub

' Creates all the controls and adds the m_data to the listbox from the m_data array
Private Sub AddRowControls()
    
    Dim attributes As labelAttributes
    attributes = m_controlAttributes.GetLabelDataAttributesFixed
    
    Dim rowFrame As MSForms.frame, currentLabel As MSForms.Label
    Dim row As Long, column As Long
    For row = 1 To m_recordsPerPage
        
        ' Create the row frame
        Set rowFrame = CreateFrameRow(row, False)
        Debug.Assert Not rowFrame Is Nothing
        
        ' Create a label for each colum of the row - skip first column as it is ID
        For column = 1 To GetDataColumnCount()
            ' Get the label settings and create the label
            Set currentLabel = CreateLabel(rowFrame _
                    , m_controlAttributes.GetLabelDataAttributesVariable _
                            (attributes, row, column, m_columnDimensions, m_multiSelect))
         Next column
        
    Next row
    
    ' Occurs if filtering and no records found
    If arrayGetRowCount(m_data) * arrayGetColumnCount(m_data) = 1 Then
        currentLabel.width = m_settings.getLabelNoRecordWidth()
    End If
    
End Sub

' PAGINATION
Private Sub AddPaginationNumberLabels()
    
    ReDim m_pageNumberEvents(1 To GetPaginationButtonCount())
    
    Dim i As Long, numberLabel As MSForms.Label
    For i = 1 To UBound(m_pageNumberEvents, 1)
        Set numberLabel = CreateLabel(m_parentFrameManager.parentFrame.Parent _
                , m_controlAttributes.GetLabelPaginationNumberAttributes(m_ID, i))
        Set m_pageNumberEvents(i) = New clsEventsPaginationNumbers
        With m_pageNumberEvents(i)
            Set .mylabel = numberLabel
            Set .parentListBox = Me
        End With
    Next i
End Sub

' 07-05-2025
Private Sub SetPaginationNumberLabelsVisible(visibleState As Boolean)
    
    If IsArrayAllocated(m_pageNumberEvents) = False Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(m_pageNumberEvents, 1)
        m_controlAccessors.GetPaginationNumberLabel(m_ID, i).Visible = visibleState
    Next i
    
End Sub

Private Sub RemovePaginationNumberLabels()
    
    If IsArrayAllocated(m_pageNumberEvents) = False Then Exit Sub
    
    Dim i As Long, numberLabel As MSForms.Label
    For i = 1 To UBound(m_pageNumberEvents, 1)
        m_parentFrameManager.parentFrame.Parent.Controls.Remove m_controlAccessors.GetPaginationNumberLabel(m_ID, i).name
        Set m_pageNumberEvents(i) = Nothing
    Next i
    
    Erase m_pageNumberEvents
End Sub

' Add the numbers to the page number labels based on the current page
Private Sub SetPaginationNumberCaptions()
    
    Dim pageButtonCount As Long: pageButtonCount = GetPaginationButtonCount()
    
    ' Keep a buffer of two page numbers after current page number
    Dim lastPageNumber As Long
    lastPageNumber = m_pageNumber + (pageButtonCount \ 2)
    
    If lastPageNumber < pageButtonCount Then
        lastPageNumber = pageButtonCount ' first page
    ElseIf lastPageNumber > m_pageCount Then
        lastPageNumber = m_pageCount ' last page
    End If

    ' Update the captions and tags of the numbers
    Dim i As Long, pageNumber As Long, currentLabel As MSForms.Label
    For i = 1 To pageButtonCount
        Set currentLabel = m_controlAccessors.GetPaginationNumberLabel(m_ID, i)
        pageNumber = lastPageNumber - pageButtonCount + i
        currentLabel.caption = pageNumber
        currentLabel.tag = pageNumber
    Next i
End Sub

' Enable/disable page buttons based on selected page
Private Sub SetPaginationNumberLabelStatus()
    Dim i As Long, currentLabel As MSForms.Label

    For i = 1 To GetPaginationButtonCount()
        Set currentLabel = m_controlAccessors.GetPaginationNumberLabel(m_ID, i)

        If CLng(currentLabel.tag) = m_pageNumber Then
            currentLabel.forecolor = m_settings.getPaginationLabelFontColorGrayed()
            currentLabel.fontUnderline = False
        Else
            currentLabel.forecolor = m_settings.getPaginationLabelFontColor()
            currentLabel.fontUnderline = m_settings.getPaginationLabelFontUnderline()
        End If
    Next i
End Sub

Private Sub AddPaginationArrowLabels()

    ReDim m_pageArrowEvents(1 To 2)
    
    Dim i As Long, numberLabel As MSForms.Label
    For i = 1 To 2
        Set numberLabel = CreateLabel(m_parentFrameManager.parentFrame.Parent _
                    , m_controlAttributes.GetLabelPaginationArrowAttributes(m_ID, i))
        Set m_pageArrowEvents(i) = New clsEventsPaginationArrows
        With m_pageArrowEvents(i)
            Set .mylabel = numberLabel
            Set .parentListBox = Me
        End With
    Next i

End Sub

' 07-05-2025
Private Sub SetPaginationArrowLabelsVisible(visibleState As Boolean)
    
    If IsArrayAllocated(m_pageNumberEvents) = False Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(m_pageArrowEvents, 1)
        m_controlAccessors.GetPaginationArrowLabel(m_ID, i).Visible = visibleState
    Next i

End Sub

Private Sub RemovePaginationArrowLabels()

    If IsArrayAllocated(m_pageArrowEvents) = False Then Exit Sub
    
    Dim i As Long, numberLabel As MSForms.Label
    For i = 1 To UBound(m_pageArrowEvents, 1)
         m_parentFrameManager.parentFrame.Parent.Controls.Remove m_controlAccessors.GetPaginationArrowLabel(m_ID, i).name
        Set m_pageArrowEvents(i) = Nothing
    Next i
    
    Erase m_pageArrowEvents

End Sub

Private Sub SetPaginationArrowLabels()
    
    Dim currentLabelLeft As MSForms.Label, currentLabelRight As MSForms.Label
    
    Set currentLabelLeft = m_controlAccessors.GetPaginationArrowLabel(m_ID, 2)
    Set currentLabelRight = m_controlAccessors.GetPaginationArrowLabel(m_ID, 1)
    
    
    currentLabelLeft.forecolor = m_settings.getPaginationLabelFontColor()
    currentLabelRight.forecolor = m_settings.getPaginationLabelFontColor()
    
    If m_pageNumber = m_pageCount Then
        currentLabelRight.forecolor = m_settings.getPaginationLabelFontColorGrayed()
    End If
    If m_pageNumber = 1 Then
        currentLabelLeft.forecolor = m_settings.getPaginationLabelFontColorGrayed()
    End If
End Sub

Private Function GetPaginationButtonCount() As Long
    Dim buttonLimit As Long: buttonLimit = m_settings.getPaginationButtonLimit()
    GetPaginationButtonCount = IIf(m_pageCount <= buttonLimit, m_pageCount, buttonLimit)
End Function

Private Sub AddMultiSelectControlsHeader()

    Dim checkboxLabel As MSForms.Label, checkboxFrame As MSForms.frame
    
    Set checkboxFrame = CreateCheckboxFrame(m_controlAccessors.GetHeaderFrame, GetCheckboxHeaderAttributes())
    Set checkboxLabel = CreateCheckboxLabel(checkboxFrame, m_controlAttributes.GetCheckboxLabelHeaderAttributes(checkboxFrame))
    
    With m_checkboxHeaderEvents
        Set .mylabel = checkboxLabel
        Set .parentListBox = Me
    End With
End Sub

Private Sub AddMultiSelectControlsRow()
    
    Dim rowFrame As MSForms.frame
    Dim checkboxLabel As MSForms.Label, checkboxFrame As MSForms.frame

    ReDim m_checkboxEvents(1 To m_recordsPerPage)
    Dim row As Long
     For row = 1 To m_recordsPerPage
        Set rowFrame = m_controlAccessors.GetRowFrame(row)  'm_parentFramemanager.parentframe.Controls(BuildRowFrameName(row))
       
        ' Create the frame and label for the checkbox
        Set checkboxFrame = CreateCheckboxFrame(rowFrame, GetCheckboxAttributes(row))
        Set checkboxLabel = CreateCheckboxLabel(checkboxFrame _
                            , m_controlAttributes.GetCheckboxLabelAttributes(checkboxFrame, row))
        
        ' Set the events
        Set m_checkboxEvents(row) = New clsEventsRowCheckbox
        With m_checkboxEvents(row)
            Set .mylabel = checkboxLabel
            Set .parentListBox = Me
        End With
         
    Next row
    
End Sub

Private Sub ShowMultiSelectControlsHeader()
    Dim show As Boolean: show = m_multiSelect <> fmMultiSelectSingle
    m_controlAccessors.GetCheckFrameHeader().Visible = show
End Sub

Private Sub ShowMultiSelectControlsRow()
    
    Dim show As Boolean: show = m_multiSelect <> fmMultiSelectSingle
    
    Dim row As Long
    For row = 1 To m_recordsPerPage
        m_controlAccessors.GetCheckFrame(row).Visible = show
    Next row
End Sub

Private Sub AddHeaderControls()
    
    ' Create the row frame
    Dim rowFrame As MSForms.frame, labelFrame As MSForms.frame
    Set rowFrame = CreateFrameRow(0, True)
    
    Dim column As Long
    For column = 1 To GetDataColumnCount()
    
        ' Create top frame - used to center the label
        Set labelFrame = CreateFrameHeaderColumn(rowFrame _
                        , m_controlAttributes.GetFrameHeaderAttributes(column, m_columnDimensions, m_multiSelect))
        Debug.Assert Not labelFrame Is Nothing
        
        ' create the label
        CreateLabel labelFrame, m_controlAttributes.GetLabelHeaderAttributes(column, m_columnDimensions)
        
        ' Create label sort icon
        CreateLabelHeaderSort labelFrame, m_controlAttributes.GetLabelHeaderSortAttributes(labelFrame.width, column)
        
    Next column
    
End Sub

' Replace the listbox data with new data. We do this after sorting
Private Sub FillControlsWithData()
    
    ' check if the controls have been created
    Dim rowFrame As MSForms.frame, currentLabel As MSForms.Label
    Dim row As Long, column As Long
    For row = 1 To getRecordCountCurrentPage
        
        ' Get the row frame
        Set rowFrame = m_controlAccessors.GetRowFrame(row)
        
        ' read through each label in the row frame. Start at 2 as first data
        Dim labelColumn As Long: labelColumn = 1
        For column = GetFirstDataColumn() To GetLastDataColumn()
            Set currentLabel = m_controlAccessors.GetRowLabel(row, labelColumn)
            currentLabel.caption = m_data(GetRecordDetailsFromPageRow(row).realRowPosition, column)
            labelColumn = labelColumn + 1
        Next column
                        
    Next row
End Sub

Private Sub FillHeaderControlsWithData()
   
   Dim rowFrame As MSForms.frame
   Set rowFrame = m_controlAccessors.GetHeaderFrame()
   
   Dim j As Long, currentLabel As MSForms.Label, labelFrame As MSForms.frame
   ' Ignore any header fields supplied past the number of columns
   For j = 1 To GetDataColumnCount()
    
        ' Create top frame - used to center the label
        Set labelFrame = m_controlAccessors.GetHeaderLabelFrame(j)
        
        ' create the label
        Set currentLabel = m_controlAccessors.GetHeaderLabel(j)
        If j > UBound(m_headerFields, 2) Then
            ReDim Preserve m_headerFields(1 To 1, 1 To j)
            m_headerFields(1, j) = "Column " & j
        End If
        currentLabel.caption = m_headerFields(1, j)
                
    Next j
End Sub

' Hide controls that do not have data - we do this after filtering
Private Sub ShowFilteredControlsOnly()

    Dim recordCount As Long: recordCount = arrayGetRowCount(m_data)
    
    Dim rowFrame As MSForms.frame, currentLabel As MSForms.Label
    Dim row As Long
    
    For row = 1 To m_recordsPerPage
        
        ' Get the row frame
        Set rowFrame = m_controlAccessors.GetRowFrame(row)
        ' If label is in row with no data then hide
        rowFrame.Visible = IIf(row <= recordCount, True, False)

    Next row
    
End Sub

' Pagination
Friend Sub LoadPageFromNumbers(labelType As String)

    Dim newPageNumber As Long: newPageNumber = CLng(labelType)
    
    ' If current page then don't do anything
    If newPageNumber = m_pageNumber Then Exit Sub
    
    m_pageNumber = newPageNumber

    Call FillControlsWithData
    Call HideUnusedRows
    Call SetPaginationNumberCaptions
    Call ResizePagination
    Call SetPaginationNumberLabelStatus
    Call SetPaginationArrowLabels
    
    Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)
    
End Sub

Friend Sub LoadPageFromArrows(labelType As String)

    If labelType = "right" Then
        If m_pageNumber = m_pageCount Then Exit Sub
        m_pageNumber = m_pageNumber + 1
    ElseIf labelType = "left" Then
        If m_pageNumber = 1 Then Exit Sub
        m_pageNumber = m_pageNumber - 1
    End If
    
    Call FillControlsWithData
    Call HideUnusedRows
    Call SetPaginationNumberCaptions
    Call SetPaginationNumberLabelStatus
    Call SetPaginationArrowLabels
    Call ResizePagination

    Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)
End Sub

' Get the number of the start and end records for the current page
Private Function GetPageRecords() As visibleRecordRange
    Dim totalRecordCount As Long: totalRecordCount = GetDataRowCount
    With GetPageRecords
        .Start = (m_recordsPerPage * (m_pageNumber - 1)) + 1
        .end = .Start + m_recordsPerPage - 1
        If .end > totalRecordCount Then .end = totalRecordCount
    End With
End Function

Private Function GetPageCount() As Long
    Dim totalRecordCount As Long: totalRecordCount = GetDataRowCount
    
    Dim fullPages As Long: fullPages = totalRecordCount \ m_recordsPerPage
    Dim notFullPage As Long: notFullPage = IIf(totalRecordCount Mod m_recordsPerPage = 0, 0, 1)

    GetPageCount = fullPages + notFullPage
End Function

Private Sub DisplayMessageInRow(Optional ByVal row As Long = 1, Optional ByVal column As Long = 1)

    Dim rowFrame As MSForms.frame, rowLabel As MSForms.Label
    
    Set rowFrame = m_controlAccessors.GetRowFrame(1)
    Set rowLabel = m_controlAccessors.GetRowLabel(1, 1)
    With rowLabel
        .Visible = True
        .caption = "No matching records found."
        .width = 150
    End With
            
End Sub

Private Sub ShowRowLabels(ByVal Visible As Boolean, Optional ByVal row As Long = 1)
    
    ' Get the row frame
    Dim rowFrame As MSForms.frame
    Set rowFrame = m_controlAccessors.GetRowFrame(1)
    
    Dim j As Long
    For j = 1 To GetDataColumnCount()
        m_controlAccessors.GetRowLabel(row, j).Visible = Visible
    Next j
End Sub

Private Sub HideUnusedRows()
    
    Dim row As Long
    For row = 1 To m_recordsPerPage
        m_controlAccessors.GetRowFrame(row).Visible _
                = IIf(row <= getRecordCountCurrentPage, True, False)
    Next row
End Sub

Private Sub HideUnusedRowsNoRecords()
    
    Dim row As Long
    For row = 2 To m_recordsPerPage
        m_controlAccessors.GetRowFrame(row).Visible = False
    Next row
End Sub

Private Sub ResizeColumns()
      
    ' check if the controls have been created
    If IsListboxBuilt() = False Then Exit Sub
    
    Dim rowFrame As MSForms.frame, Label As MSForms.Label
    Dim row As Long, column As Long
    For row = 1 To m_recordsPerPage
        ' Get the row frame
        Set rowFrame = m_controlAccessors.GetRowFrame(row)
        For column = 1 To GetDataColumnCount()
            Debug.Assert Not rowFrame Is Nothing
             Set Label = m_controlAccessors.GetRowLabel(row, column)
            Label.left = m_calcs.GetColumnLeft(column, m_columnDimensions, m_multiSelect)
            Label.width = m_calcs.GetColumnWidth(column, m_columnDimensions)
        Next column
    Next row
    
End Sub

Private Sub ResizeHeaders()
    
    If IsListboxBuilt() = False Then Exit Sub
    Dim row As Long: row = 0
    
    Dim rowFrame As MSForms.frame
    Set rowFrame = m_controlAccessors.GetHeaderFrame()
    
    Dim labelFrame As MSForms.frame
    Dim labelValue As String, labelWidth As Long, labelLeft As Long
    Dim j As Long, columnLabel As MSForms.Label, sortLabel As MSForms.Label
   
    Dim frameAttributes As frameHeaderAttributes, labelHeaderAttributes As labelAttributes
    For j = 1 To GetDataColumnCount()
        Debug.Assert Not rowFrame Is Nothing
        
        ' Resize column frame
        Set labelFrame = m_controlAccessors.GetHeaderLabelFrame(j)
        frameAttributes = m_controlAttributes.GetFrameHeaderAttributes(j, m_columnDimensions, m_multiSelect)
        labelFrame.left = frameAttributes.left
        labelFrame.width = frameAttributes.width
        
        ' Resize column label
        Set columnLabel = m_controlAccessors.GetHeaderLabel(j)
        labelHeaderAttributes = m_controlAttributes.GetLabelHeaderAttributes(j, m_columnDimensions)
        columnLabel.width = labelHeaderAttributes.width
        columnLabel.left = labelHeaderAttributes.left

        ' Resize column sort symbol
        Set sortLabel = m_controlAccessors.GetHeaderLabelSymbol(j)
        With m_controlAttributes.GetLabelHeaderSortAttributes(labelFrame.width, j)
            sortLabel.width = .width
            sortLabel.left = .left
        End With
    Next j
    
End Sub

' Resize the Listbox height based on the number of records
Private Sub ResizeHeight()
    If m_AutomaticHeight = False Then Exit Sub
    If IsListboxBuilt() = False Then Exit Sub
    m_parentFrameManager.parentFrame.Height = m_calcs.GetMainFrameHeight(m_recordsPerPage)
End Sub

' Resize the Listbox width based on the number of columns
Private Sub ResizeWidth()
    If m_AutomaticWidth = False Then Exit Sub
    If IsListboxBuilt() = False Then Exit Sub
    m_parentFrameManager.parentFrame.width = m_calcs.GetMainFrameWidth(GetDataColumnCount(), m_columnDimensions, m_multiSelect)
End Sub

Private Sub SetScrollBars()
    If IsListboxBuilt() = False Then Exit Sub
    ' Resize parent frame
    With m_parentFrameManager.parentFrame
        .ScrollWidth = m_calcs.GetRowFrameWidth(GetDataColumnCount(), m_columnDimensions, m_multiSelect)
        .ScrollHeight = m_calcs.GetMainFrameHeight(m_recordsPerPage)
    End With

End Sub

Private Sub ResizeSearch()
    If m_controlAccessors.CheckSearchBoxExists(m_ID) = False Then Exit Sub

    Dim searchTextbox As MSForms.TextBox
    Set searchTextbox = m_controlAccessors.GetSearchBox(m_ID)

    With searchTextbox
        .left = m_calcs.GetSearchBoxLeft(m_parentFrameManager.parentFrame.left, m_parentFrameManager.parentFrame.width)
        .width = m_settings.getSearchBoxWidth()
    End With
    
    Dim searchLabel As MSForms.Label
    Set searchLabel = m_controlAccessors.GetSearchLabel(m_ID)
    With searchLabel
        .left = m_calcs.GetSearchBoxLeft(m_parentFrameManager.parentFrame.left, m_parentFrameManager.parentFrame.width) - m_settings.getsearchBoxLabelWidth()
        .width = m_settings.getsearchBoxLabelWidth()
    End With
End Sub

Private Sub ResizePagination()
    
    Dim pageButtonCount As Long: pageButtonCount = GetPaginationButtonCount()
    Dim i As Long, currentLabel As MSForms.Label
    Dim previousLeft As Long, digitCount As Long, labelWidth As Long
    For i = pageButtonCount To 1 Step -1

        Set currentLabel = m_controlAccessors.GetPaginationNumberLabel(m_ID, i)
        digitCount = Len(currentLabel.tag)
         labelWidth = m_calcs.GetPaginationLabelNumberWidth(digitCount)
        If i = pageButtonCount Then
            currentLabel.left = m_calcs.GetPaginationLabelNumberLeft(digitCount)
        Else
            currentLabel.left = previousLeft - labelWidth
        End If
        previousLeft = currentLabel.left
        currentLabel.width = labelWidth
    Next i
    
    For i = 1 To 2
        Set currentLabel = m_controlAccessors.GetPaginationArrowLabel(m_ID, i)
        currentLabel.left = m_calcs.GetPaginationLabelArrowLeft(i)
    Next i
End Sub

' CREATE THE LISTBOX CONTROLS
Private Sub RemoveSearch()
    If IsListboxBuilt() = False Then Exit Sub
    m_parentFrameManager.parentFrame.Parent.Controls.Remove m_controlAccessors.GetSearchBox(m_ID).name
    m_parentFrameManager.parentFrame.Parent.Controls.Remove m_controlAccessors.GetSearchLabel(m_ID).name
End Sub

Private Sub CreateSearch()
    Call CreateSearchTextBox
    Call CreateSearchBoxLabel
End Sub


Private Sub CreateSearchTextBox()


    Dim searchBox As MSForms.TextBox
    Set searchBox _
        = m_parentFrameManager.parentFrame.Parent.Controls.Add("Forms.Textbox.1", m_nameHelper.BuildSearchTextboxName(m_ID), True)
    
    With searchBox
        .tag = -1

        .BorderStyle = fmBorderStyleSingle
        .SelectionMargin = False
        .MultiLine = True
        .BorderColor =  RGB(211, 211, 211)

        .left = m_calcs.GetSearchBoxLeft(m_parentFrameManager.parentFrame.left, m_parentFrameManager.parentFrame.width)
        .width = m_settings.getSearchBoxWidth()
        .Height = m_settings.getSearchBoxHeight()
        .top = m_calcs.GetSearchTop(m_parentFrameManager.parentFrame.top)
        .Font.Size = m_settings.searchBoxFontSize
        .Font.name = m_settings.searchBoxFontName
   
        Set m_searchTextboxEvents(1).myTextbox = searchBox
        Set m_searchTextboxEvents(1).parentListBox = Me
    End With

    ' TxtColor forecolor:="#43545F", _
    ' EnterColor:="#005ea2", _
    ' TitleColor:="#ababab", _
    ' AlertColor:="#f72111", _
    ' SuccessColor:="#22ab2b"
    ' tbox.clasBox Me, "Border"
    ' cBtn.classButton Me
End Sub



Private Sub CreateSearchBoxLabel()
    
    Dim searchboxLabel As MSForms.Label
    Set searchboxLabel = m_parentFrameManager.parentFrame.Parent.Controls.Add("Forms.Label.1", m_nameHelper.BuildSearchLabelName(m_ID), True)
    With searchboxLabel
        .caption = "Search: "
        .tag = -1
        .left = m_calcs.GetSearchBoxLeft(m_parentFrameManager.parentFrame.left, m_parentFrameManager.parentFrame.width) - m_settings.getsearchBoxLabelWidth()
        .width = m_settings.getsearchBoxLabelWidth()
        .top = m_calcs.GetSearchTop(m_parentFrameManager.parentFrame.top)
        .Height = m_settings.getSearchBoxHeight()
        .Font.Size = m_settings.searchBoxLabelFontSize
        .Font.name = m_settings.searchBoxLabelFontName
        .forecolor = rgbGray
        Set m_searchLabelEvents(1).mylabel = searchboxLabel
        Set m_searchLabelEvents(1).parentListBox = Me
#If ShowLabels = 1 Then
    .BorderStyle = fmBorderStyleSingle
    .SpecialEffect = fmSpecialEffectEtched ' set to see frame
#End If

    End With
    
End Sub

Private Function CreateImage(ByVal filename As String _
                    , ByVal controlName As String _
                    , ByVal parentFrame As MSForms.frame _
                    , Optional ByVal isVisible As Boolean = True) As MSForms.Image
    If parentFrame Is Nothing Then
        Err.Raise errObjectNotSet, "CreateImage" _
            , "m_ParentFrame should be set to the UserForm frame that will contain the listbox"
    End If
    
    Set CreateImage = parentFrame.Controls.Add("Forms.Image.1", controlName, isVisible)
    With CreateImage
        .Picture = LoadPicture(ThisWorkbook.Path & Application.PathSeparator & filename)
        .PictureAlignment = fmPictureAlignmentCenter
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
    End With
    
End Function
'

Private Function CreateCheckboxFrame(parentFrame As MSForms.frame _
                                    , attributes As checkboxAttributes) As MSForms.frame
    Dim currentCheckbox As MSForms.frame
    Set currentCheckbox = parentFrame.Controls.Add("Forms.Frame.1", attributes.name, True)
    With currentCheckbox
        .left = attributes.left
        .Height = attributes.Height
        .top = attributes.top
        .width = attributes.width
        .caption = attributes.caption 'ChrW(&HE03B)
        .SpecialEffect = fmButtonEffectSunken
        .borderStyle = fmBorderStyleSingle
        .borderColor = rgb(199, 196, 196)
        .tag = attributes.tag
    End With
    Set CreateCheckboxFrame = currentCheckbox
    Debug.Assert Not CreateCheckboxFrame Is Nothing
End Function

Private Function CreateCheckboxLabel(parentFrame As MSForms.frame _
                , attributes As labelAttributes) As MSForms.Label
                
    Dim checkboxLabel As MSForms.Label
    Set checkboxLabel = parentFrame.Controls.Add("Forms.Label.1", attributes.name, True)
    With checkboxLabel
        .fontName = attributes.fontName
        .FontBold = attributes.FontBold
        .fontSize = attributes.fontSize
        .caption = attributes.caption 'ChrW(&HE03B) 
        .Height = attributes.Height
        .tag = attributes.tag
        .width = attributes.width
        .forecolor = rgb(2, 78, 139)
    End With
    Set CreateCheckboxLabel = checkboxLabel
    Debug.Assert Not CreateCheckboxLabel Is Nothing
End Function

Private Function CreateLabel(parentFrame As Variant _
                            , attributes As labelAttributes) As MSForms.Label
    ' Create label control
    Dim currentLabel As MSForms.Label
    Set currentLabel = parentFrame.Controls.Add("Forms.Label.1", attributes.name, True)
    With currentLabel
        .caption = attributes.caption
        .textAlign = attributes.textAlign
        .left = attributes.left
        .width = attributes.width
        .Height = attributes.Height
        .top = attributes.top
        .backcolor = attributes.backcolor
        .forecolor = attributes.forecolor
        .Font.name = attributes.fontName
        .fontUnderline = attributes.fontUnderline
        .Font.Size = attributes.fontSize
        .tag = attributes.tag
        
#If ShowLabels = 1 Then
        .BorderStyle = fmBorderStyleSingle ' turn on to see borders
#End If
    End With
    
    ' Return the label
    Set CreateLabel = currentLabel
    
    Debug.Assert Not CreateLabel Is Nothing

End Function

' Creates the label with the header sort
Private Function CreateLabelHeaderSort(parentFrame As MSForms.frame _
                            , attributes As labelAttributes) As MSForms.Label

    ' Create label control
    Dim currentLabel As MSForms.Label
    Set currentLabel = parentFrame.Controls.Add("Forms.Label.1", attributes.name, True)
    With currentLabel
        .Font.name = attributes.fontName
        .Font.Size = attributes.fontSize
        .caption = attributes.caption
        .textAlign = attributes.textAlign
        .left = attributes.left
        .width = attributes.width
        .Height = attributes.Height
        .top = attributes.top
        .backcolor = attributes.backcolor
        .forecolor = attributes.forecolor
        .tag = attributes.tag
#If ShowLabels = 1 Then
        .BorderStyle = fmBorderStyleSingle ' turn on to see borders
#End If
    End With
    
    
    ' Return the label
    Set CreateLabelHeaderSort = currentLabel
    
    Debug.Assert Not CreateLabelHeaderSort Is Nothing
    Debug.Assert Trim(CreateLabelHeaderSort.name) <> ""

End Function

Private Function CreateFrameRow(ByVal row As Long _
                                , isHeaderRow As Boolean) As MSForms.frame
                                
    If m_parentFrameManager.parentFrame Is Nothing Then
        Err.Raise errObjectNotSet, "CreateFrame" _
            , "m_parentFrameManager.parentFrame should be set to the UserForm frame that will contain the listbox"
    End If
    
    Dim name As String: name _
        = IIf(isHeaderRow, m_nameHelper.BuildHeaderFrameName(), m_nameHelper.BuildRowFrameName(row))
    Set CreateFrameRow = m_parentFrameManager.parentFrame.Controls.Add("Forms.Frame.1", name, True) ' zz

    With CreateFrameRow
        .caption = ""
        .left = -1
        .Height = IIf(isHeaderRow, m_settings.getColumnHeaderRowHeight(), m_settings.getDataRowHeight())
        .width = m_calcs.GetRowFrameWidth(GetDataColumnCount(), m_columnDimensions, m_multiSelect)
        .top = IIf(isHeaderRow, m_calcs.GetHeaderTop(), m_calcs.GetDataRowTop(row))
        .tag = row
        .backcolor = IIf(isHeaderRow, m_settings.columnHeaderBackgroundColor(), m_settings.getRowBackgroundColor())
#If ShowLabels = 1 Then
    .SpecialEffect = fmSpecialEffectEtched ' set to see frame
#End If
    End With
    
End Function

Private Function CreateFrameHeaderColumn(rowFrame As MSForms.frame _
                                            , attributes As frameHeaderAttributes) As MSForms.frame
                                
    If m_parentFrameManager.parentFrame Is Nothing Then
        Err.Raise errObjectNotSet, "CreateFrame" _
            , "m_parentFrameManager.parentFrame should be set to the UserForm frame that will contain the listbox"
    End If
    
    Dim labelFrame As MSForms.frame
    Set labelFrame = rowFrame.Controls.Add("Forms.Frame.1", attributes.name, True)
    With labelFrame
            .left = attributes.left
            .width = attributes.width
            .top = attributes.top
            .Height = attributes.Height
            .SpecialEffect = attributes.SpecialEffect
#If ShowLabels = 1 Then
    .SpecialEffect = fmSpecialEffectEtched ' set to see frame
#End If
    End With
    
    Set CreateFrameHeaderColumn = labelFrame
    
End Function

' REMOVE CONTROLS
Private Sub RemoveHeaderControls()
    On Error Resume Next ' header controls may not exist
    ' If code stops here then make sure "Break on all errors" is not turned on in options
    m_parentFrameManager.parentFrame.Controls.Remove m_nameHelper.BuildHeaderFrameName()
End Sub

Private Sub RemoveAllDataControls()
    Dim currentControl As MSForms.Control
    For Each currentControl In m_parentFrameManager.parentFrame.Controls
        If TypeName(currentControl) = "Frame" _
                    And m_nameHelper.GetControlTypename(currentControl.name) = "rowFrame" Then
            If m_nameHelper.GetRowFromFrameName(currentControl.name) > 0 Then
                m_parentFrameManager.parentFrame.Controls.Remove currentControl.name
            End If
        End If
    Next
End Sub

' CREATE HOVER EVENTS
Private Sub CreateHeaderControlHoverEvents()

    ReDim m_headerLabelEvents(1 To GetDataColumnCount())
    ReDim m_sortEvents(1 To GetDataColumnCount())

    Dim currentLabel As MSForms.Label, labelNumber As Long, rowFrame As MSForms.frame

    Dim column As Long
    
    Set rowFrame = m_parentFrameManager.parentFrame.Controls(m_nameHelper.BuildHeaderFrameName())
    With m_frameRowHeaderEvents
        Set .myFrame = rowFrame
        Set .parentListBox = Me
    End With

    For column = 1 To GetDataColumnCount()
        Debug.Assert Not rowFrame Is Nothing
        
        Set m_headerLabelEvents(column) = New clsEventsHeaderLabel
        With m_headerLabelEvents(column)
            Set currentLabel = m_controlAccessors.GetHeaderLabel(column)
            Set .mylabel = currentLabel
            Set .parentListBox = Me
        End With

        Set m_sortEvents(column) = New clsEventsHeaderSymbol
        With m_sortEvents(column)
            Set currentLabel = m_controlAccessors.GetHeaderLabelSymbol(column)
            Set .mylabel = currentLabel
            Set .parentListBox = Me
        End With
    Next column

End Sub

Private Sub CreateDataControlHoverEvents()
        
    ReDim m_labelRowEvents((m_recordsPerPage + 1) * arrayGetColumnCount(m_data))
    ReDim m_frameRowEvents(1 To m_recordsPerPage)


    Dim currentLabel As MSForms.Label, labelNumber As Long
    Dim rowFrame As MSForms.frame, checkboxFrame As MSForms.frame ', checkboxLabel As MSForms.Label
    Dim row As Long, column As Long, labelColumn As Long
    
    For row = 1 To m_recordsPerPage
        
        ' Create events for the row frames
        Set rowFrame = m_parentFrameManager.parentFrame.Controls(m_nameHelper.BuildRowFrameName(row))

        Set m_frameRowEvents(row) = New clsEventsRowFrame
        With m_frameRowEvents(row)
            Set .myFrame = rowFrame
            Set .parentListBox = Me
        End With

        labelColumn = 1
        For column = GetFirstDataColumn To UBound(m_data, 2)
            
            labelNumber = m_nameHelper.GetLabelEventNumber(row, column, arrayGetColumnCount(m_data))
                        
            ' Create an event for each label
            Set m_labelRowEvents(labelNumber) = New clsEventsRowLabel
            With m_labelRowEvents(labelNumber)
                Set currentLabel = m_controlAccessors.GetRowLabel(row, labelColumn)
                Set .mylabel = currentLabel
                Set .parentListBox = Me
            End With
            labelColumn = labelColumn + 1
        Next column
    Next row
End Sub

' FILTERING
Private Function FilterData(ByRef text As String) As Long
    
    m_filterString = text
    Dim temp As Variant
    Call arraySetSize(temp, m_data)
    
    Dim i As Long, j As Long, row As Long
    For i = LBound(m_data, 1) To UBound(m_data, 1)
        
        For j = GetFirstDataColumn To UBound(m_data, 2)
            If InStr(1, m_data(i, j), m_filterString, vbTextCompare) > 0 Then
                row = row + 1
                Call arrayCopyRow(temp, row, m_data, i)
                Exit For
            End If
        Next j
    Next i
    
    ' Check if data changed
    If row > 0 Then
        ' Copy filtered data to the main data array
        Call arraySetSize(m_data, temp, row)
        Call arrayCopy(m_data, temp, row)
    End If
    FilterData = row
End Function

' HOVER HIGHLIGHING

Friend Sub SetRowHover(row As Long)
    If m_hoverOn = False Then Exit Sub
    If m_highlightTracker.IsRowHighlighted(GetRecordDetailsFromPageRow(row).ID) = True Then Exit Sub
    Call m_hoverRow.SetRowHover(row)
    Call m_hoverHeader.ClearHover
End Sub

Public Sub ClearHover()
    If IsListboxBuilt() Then
        m_hoverRow.ClearHover
        m_hoverHeader.ClearHover
    End If
End Sub

Friend Sub SetHeaderHoverHighlight(ByVal Label As MSForms.Label)
    If m_hoverOn = False Then Exit Sub
    Call m_hoverHeader.SetHover(m_nameHelper.GetColumnFromHeaderLabelName(Label.name))
    Call m_hoverRow.ClearHover
End Sub

Friend Sub DoubleClick(row As Long)
    RaiseEvent ItemDoubleClicked(row)
End Sub

' SELECTION HIGHLIGHING

' Used when the user selects a row
Friend Sub SelectRow(row As Long)

    Dim details As recordDetails
    details = GetRecordDetailsFromPageRow(row)

    Call m_hoverRow.ClearHover
    Call m_highlightTracker.SelectRow(details, m_multiSelect)
    Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)
    
    If m_highlightTracker.SelectedRowCount > 0 Then
        RaiseEvent ItemSelected(details.realRowPosition)
    Else
        RaiseEvent ItemSelected(details.realRowPosition)
    End If
End Sub

' Use to set a pageRow true or false - when user used the SetSelected method
Private Sub SetSelectRow(index As Long, turnOn As Boolean)
    
    Call m_hoverRow.ClearHover
    Call m_highlightTracker.SetRow(GetRecordDetailsFromIndex(index), m_multiSelect, turnOn)
    Call m_highlightRow.ApplyHighlight(GetPageRecords(), m_highlightTracker.SelectedRowNumbers)

    If m_highlightTracker.SelectedRowCount > 0 Then
        RaiseEvent ItemSelected(index)
    Else
        RaiseEvent ItemDeselected(index)
    End If
End Sub


Private Function GetRecordDetailsFromPageRow(row As Long) As recordDetails
    With GetRecordDetailsFromPageRow
       .realRowPosition = GetRealRowFromPageRow(row)
       .ID = CLng(m_data(.realRowPosition, 1))
    End With
End Function

Private Function GetRecordDetailsFromIndex(index As Long) As recordDetails
    With GetRecordDetailsFromIndex
       .realRowPosition = index
       .ID = CLng(m_data(index, 1))
    End With
End Function

' SORTING

Private Sub SetColumnSortStatus()
    Call m_sortTracker.SetColumnSortStatus(GetDataColumnCount())
End Sub

Private Sub SortData(ByVal column As Long)

    Dim sortDirection As Long: sortDirection = m_sortTracker.SetSortDirection(column)
    
    Dim sortColumn As Long: sortColumn = (GetFirstDataColumn() - 1) + column
    
    m_data = WorksheetFunction.Sort(m_data, sortColumn, sortDirection)

End Sub

Private Sub SetSortSymbol(ByVal column As Long)
    Debug.Assert column > 0 And column <= UBound(m_data, 2)
    Dim j As Long
    For j = 1 To GetDataColumnCount()
        m_controlAccessors.GetHeaderLabelSymbol(j).caption = m_settings.GetLabelSortSymbol(soNeutral)
    Next j
    
    m_controlAccessors.GetHeaderLabelSymbol(column).caption = _
        m_settings.GetLabelSortSymbol(m_sortTracker.GetSortDirection(column))
                    
End Sub

Private Function GetCheckboxAttributes(row As Long) As checkboxAttributes
    
    With GetCheckboxAttributes
        .name = m_nameHelper.BuildCheckboxName(row)
        .tag = row
    
        .left = m_calcs.GetCheckboxLeft()
        .width = m_settings.getCheckboxWidth()
        .top = m_settings.getCheckboxTop()
        .Height = m_settings.getCheckboxHeight()
        
    End With
    
End Function

Private Function GetCheckboxHeaderAttributes() As checkboxAttributes
    
    With GetCheckboxHeaderAttributes
        .name = m_nameHelper.BuildCheckboxNameHeader
        .tag = 0
    
        .left = m_calcs.GetCheckboxLeft()
        .width = m_settings.getCheckboxWidth()
        .top = m_settings.getCheckboxTop()
        .Height = m_settings.getCheckboxHeight()
        
    End With
    
End Function

' HELPER FUNCTIONS
Private Function IsArrayEventsRowLabelAllocated(ByRef arr() As clsEventsRowLabel) As Boolean
    IsArrayEventsRowLabelAllocated = UBound(arr, 1) <> -1
End Function

Private Function IsDataFiltered() As Boolean
    IsDataFiltered = Not (arrayGetRowCount(m_data) = arrayGetRowCount(m_dataStore))
End Function

Private Function IsListboxBuilt() As Boolean
'    IsListboxBuilt = m_parentFramemanager.parentframe.Controls.Count > 1
    IsListboxBuilt = m_parentFrameManager.parentFrame.Controls.Count > 1
End Function

Private Function IsMultiSelect() As Boolean
    IsMultiSelect = m_multiSelect <> fmMultiSelectSingle
End Function

Private Sub RecallData()
    m_data = m_dataStore
End Sub

Private Sub BackupData()
    m_dataStore = m_data
End Sub

Private Function GetFirstDataColumn() As Long
    GetFirstDataColumn = LBound(m_data, 2) + 1
End Function

Private Function GetLastDataColumn() As Long
    GetLastDataColumn = UBound(m_data, 2)
End Function

Private Function GetDataColumnCount() As Long
    GetDataColumnCount = arrayGetColumnCount(m_data) - 1
End Function

Private Function GetDataRowCount() As Long
    GetDataRowCount = arrayGetRowCount(m_data)
End Function

Private Function GetDataRow(row As Long) As Variant
    ReDim newRow(1 To 1, 1 To UBound(m_data, 2) - 1) As Variant
    Dim i As Long
    For i = 2 To UBound(m_data, 2)
        newRow(1, i - 1) = m_data(row, i)
    Next i
    GetDataRow = newRow
End Function

Private Function GetPageRowFromRealRow(realRow As Long)
    GetPageRowFromRealRow = realRow Mod m_recordsPerPage
    GetPageRowFromRealRow = IIf(GetPageRowFromRealRow = 0, m_recordsPerPage, GetPageRowFromRealRow)
    Debug.Assert GetPageRowFromRealRow > 0
End Function

Private Function GetRealRowFromPageRow(pageRow As Long)
    GetRealRowFromPageRow = pageRow + ((m_pageNumber - 1) * m_recordsPerPage)
    Debug.Assert GetRealRowFromPageRow > 0
End Function

Private Function getRecordCountCurrentPage() As Long
    
    Dim totalRecordCount As Long: totalRecordCount = GetDataRowCount
    
    getRecordCountCurrentPage = m_recordsPerPage
    If m_pageNumber = m_pageCount Then
        ' If last page get number of records on it
        getRecordCountCurrentPage = totalRecordCount Mod m_recordsPerPage
        ' check for case where the last page is full
        If getRecordCountCurrentPage = 0 Then getRecordCountCurrentPage = m_recordsPerPage
    End If
    
    Debug.Assert getRecordCountCurrentPage > 0
    
End Function

Private Function IsColumnDimensionArrayAllocated(ByRef arr() As ColumnDimension) As Boolean
    IsColumnDimensionArrayAllocated = (Not Not arr) <> 0
End Function

Private Function GetColumnDimensions(arr() As ColumnDimension) As Long
    GetColumnDimensions = UBound(arr, 1) - LBound(arr, 1) + 1
End Function

Private Sub SetColumnDimensionDefaults()
    
    Dim ColumnCount As Long
    ColumnCount = GetDataColumnCount()
    
    ' Check if the column dimsions are not initialized
    Dim existingColumnCount As Long, leftPosition As Long
    If IsColumnDimensionArrayAllocated(m_columnDimensions) = False Then
        ReDim m_columnDimensions(1 To ColumnCount)
        existingColumnCount = 0
        leftPosition = m_settings.getFrameLeftPadding()
    Else
        ' Get the number of column dimensions
        existingColumnCount = GetColumnDimensions(m_columnDimensions)
        ' If dimensions for each column already then exit
        If existingColumnCount >= ColumnCount Then Exit Sub
        
        ' Resize the array to include the new dimensions
        ReDim Preserve m_columnDimensions(1 To ColumnCount)
        leftPosition = m_columnDimensions(existingColumnCount).left _
                        + m_columnDimensions(existingColumnCount).width
    End If
    
    ' Copy the default widths to the new positions
    Dim i As Long
    For i = existingColumnCount + 1 To UBound(m_columnDimensions)
        m_columnDimensions(i).width = m_settings.getLabelWidthDefault
        m_columnDimensions(i).left = leftPosition
        leftPosition = leftPosition + m_settings.getLabelWidthDefault
    Next i
  
End Sub

' Set the column dimensions for given widths
Private Sub SetColumnDimensions(ByVal columnWidths As String)
    
    Dim ColumnCount As Long
    ColumnCount = GetDataColumnCount()
    
    Dim widths As Variant
    widths = Split(columnWidths, ";")
       
    If ColumnCount = -1 Then ColumnCount = UBound(widths) - LBound(widths) + 1
    ReDim m_columnDimensions(1 To ColumnCount)
    
    Dim widthsColumnCount As Long
    widthsColumnCount = IIf(arrayGetRowCount(widths) > ColumnCount, ColumnCount, arrayGetRowCount(widths))
    
    ' Copy the given dimensions to the column dimension array
    Dim leftPosition As Long: leftPosition = m_settings.getFrameLeftPadding()
    Dim i As Long

    For i = 1 To ColumnCount
        If i <= widthsColumnCount Then
            m_columnDimensions(i).width = widths(i - 1)
            m_columnDimensions(i).left = leftPosition
            leftPosition = leftPosition + widths(i - 1)
        Else
            m_columnDimensions(i).width = m_settings.getLabelWidthDefault
            m_columnDimensions(i).left = leftPosition
            leftPosition = leftPosition + m_settings.getLabelWidthDefault
        End If
    Next i
      
End Sub

Private Function IsColumnWidthStringValid(columnWidths) As Boolean
    
    Dim width As Variant
    For Each width In Split(columnWidths, ";")
        On Error Resume Next
        width = CLng(width)
        If IsNumeric(width) = False Or InStr(width, ",") > 0 Then
            Exit Function
        End If
    Next
    IsColumnWidthStringValid = True
    
End Function

' Public properties for fonts etc.
Public Property Let dataFontColor(ByVal argdataFontColor As Long)
     m_settings.dataFontColor = argdataFontColor
End Property
Public Property Let dataFontName(ByVal argdataFontName As String)
     m_settings.dataFontName = argdataFontName
End Property
Public Property Let dataFontSize(ByVal argdataFontSize As Long)
     m_settings.dataFontSize = argdataFontSize
End Property
Public Property Let columnHeaderBackgroundColor(ByVal argcolumnHeaderBackgroundColor As Long)
     m_settings.columnHeaderBackgroundColor = argcolumnHeaderBackgroundColor
End Property
Public Property Let columnHeaderFontColor(ByVal argcolumnHeaderFontColor As Long)
     m_settings.columnHeaderFontColor = argcolumnHeaderFontColor
End Property
Public Property Let columnHeaderFontName(ByVal argcolumnHeaderFontName As String)
     m_settings.columnHeaderFontName = argcolumnHeaderFontName
End Property
Public Property Let columnHeaderFontSize(ByVal argcolumnHeaderFontSize As Long)
     m_settings.columnHeaderFontSize = argcolumnHeaderFontSize
End Property
Public Property Let rowSelectionColor(ByVal argrowSelectionColor As Long)
     m_settings.rowSelectionColor = argrowSelectionColor
End Property
Public Property Let hoverColorRow(ByVal arghoverColorRow As Long)
     m_settings.hoverColorRow = arghoverColorRow
End Property
Public Property Let hoverColorheader(ByVal arghoverColorheader As Long)
     m_settings.hoverColorheader = arghoverColorheader
End Property
Public Property Let listboxBackgroundColor(ByVal arglistboxBackgroundColor As Long)
     m_settings.listboxBackgroundColor = arglistboxBackgroundColor
End Property

Public Property Let searchBoxLabelFontName(ByVal argsearchBoxLabelFontName As String)
     m_settings.searchBoxLabelFontName = argsearchBoxLabelFontName
End Property
Public Property Let searchBoxLabelFontSize(ByVal argsearchBoxLabelFontSize As Long)
     m_settings.searchBoxLabelFontSize = argsearchBoxLabelFontSize
End Property
Public Property Let searchBoxFontName(ByVal argsearchBoxFontName As String)
     m_settings.searchBoxFontName = argsearchBoxFontName
End Property
Public Property Let searchBoxFontSize(ByVal argsearchBoxFontSize As Long)
     m_settings.searchBoxFontSize = argsearchBoxFontSize
End Property


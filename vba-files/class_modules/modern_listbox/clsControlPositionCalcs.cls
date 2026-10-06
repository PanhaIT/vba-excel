VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlPositionCalcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Calculation procedures for positioning controls that are part of the Listbox

Private m_settings As clsListBoxSettings
'Private m_parentFrame As frame
Private m_parentFrame As clsParentFrameManager
Private m_InitializedCorrectly As Boolean

Private Sub Class_Initialize()
    m_InitializedCorrectly = False
End Sub

Public Property Let InitializedCorrectly(newInitializedCorrectly As Boolean)
    m_InitializedCorrectly = True
End Property

Public Property Set settings(argSettings As clsListBoxSettings)
    Set m_settings = argSettings
End Property

Public Property Set parentFrame(parentFrame As clsParentFrameManager)
    Set m_parentFrame = parentFrame
End Property

' Used to ensure that parent frame was initialized
Private Function GetParentFrame() As frame
    If m_InitializedCorrectly = False Then
        Err.Raise errObjectNotSet, "clsControlPositionCalcs.GetParentFrame" _
    , "clsControlPositionCalcs should be instantiated using CreateControlPositionCalcs() " _
        & "so that the parent frame can be set."
    End If
    Set GetParentFrame = m_parentFrame.parentFrame
End Function

Public Function GetColumnLeft(column As Long _
                    , columnDimensions() As ColumnDimension _
                    , multiSelect As fmMultiSelect) As Long
    
    GetColumnLeft = m_settings.GetSelectWidth(multiSelect) _
                    + columnDimensions(column).left + m_settings.getDataLabelPaddingLeft()
End Function

Public Function GetColumnWidth(column As Long, columnDimensions() As ColumnDimension) As Long
    GetColumnWidth = columnDimensions(column).width - m_settings.getDataLabelPaddingLeft()
End Function

Public Function GetCheckboxLeft() As Long
    GetCheckboxLeft = m_settings.GetCheckboxLeft() + m_settings.getDataLabelPaddingLeft()
End Function

Public Function GetHeaderTop() As Long
    GetHeaderTop = m_settings.getframeTopPadding()

End Function
Public Function GetDataRowTop(ByVal row As Long) As Long
    With m_settings
        GetDataRowTop = GetHeaderTop + .getColumnHeaderRowHeight() + (((row - 1) * .getDataRowHeight) - row)
    End With
End Function
Public Function GetSearchBoxLeft(frameLeft As Long, frameWidth As Long) As Long
    GetSearchBoxLeft = frameLeft + frameWidth _
            - (m_settings.getSearchBoxWidth() + m_settings.getSearchBoxPaddingRight())
End Function
Public Function GetSearchTop(frameTop As Long) As Long
    GetSearchTop = frameTop - (m_settings.getSearchBoxHeight() + m_settings.getSearchBoxPaddingBottom())
End Function

Public Function GetSymbolLabelLeft(parentFrameWidth As Long) As Long
    GetSymbolLabelLeft = parentFrameWidth _
                - (m_settings.GetLabelSortSymbolWidth() + m_settings.GetLabelSortSymbolPaddingRight())

    Debug.Assert GetSymbolLabelLeft > 0
End Function

' Used for the rightmost number only
Public Function GetPaginationLabelNumberLeft(digitCount As Long) As Long
    
    Dim labelWidth As Long: labelWidth = GetPaginationLabelNumberWidth(digitCount)

    Dim arrowsWidth As Long: arrowsWidth = 2 * m_settings.getPaginationLabelWidth()
    GetPaginationLabelNumberLeft = ((GetParentFrame().left + GetParentFrame().width) _
                                - m_settings.getPaginationPaddingRight()) _
                        - labelWidth - arrowsWidth
End Function

Public Function GetPaginationLabelArrowLeft(position As Long) As Long
    GetPaginationLabelArrowLeft = ((GetParentFrame().left + GetParentFrame().width) _
                        - m_settings.getPaginationPaddingRight()) _
                        - (m_settings.getPaginationLabelWidth() * (position))

End Function

Public Function GetPaginationLabelNumberWidth(digitCount As Long) As Long
    GetPaginationLabelNumberWidth = m_settings.getPaginationLabelWidth() _
                + ((m_settings.getPaginationLabelWidth() * (digitCount - 1)) \ 2)
End Function

Public Function GetPaginationTop()
    GetPaginationTop = GetParentFrame().top + GetParentFrame().Height
End Function

Public Function GetTotalColumnWidthNew(ColumnCount As Long, columnDimensions() As ColumnDimension) As Long
    Dim i As Long
    For i = 1 To ColumnCount
        GetTotalColumnWidthNew = GetTotalColumnWidthNew + columnDimensions(i).width
    Next i
End Function

Public Function GetMainFrameHeight(recordsPerPage As Long) As Long
    GetMainFrameHeight = GetHeightAllRows(recordsPerPage) _
                                + m_settings.getframeTopPadding() _
                                + m_settings.getframeBottomPadding()
End Function

Public Function GetMainFrameWidth(ColumnCount As Long _
                                        , columnDimensions() As ColumnDimension _
                                        , multiSelect As fmMultiSelect) As Long
    GetMainFrameWidth = GetRowFrameWidth(ColumnCount, columnDimensions, multiSelect) _
                                + m_settings.getFrameLeftPadding() _
                                + m_settings.getFrameRightPadding() _

End Function

Private Function GetHeightAllRows(recordsPerPage As Long) As Long
    GetHeightAllRows = m_settings.getColumnHeaderRowHeight() _
                + (m_settings.getDataRowHeight()) * recordsPerPage
End Function

Public Function GetRowFrameWidth(ColumnCount As Long _
                    , columnDimensions() As ColumnDimension _
                    , multiSelect As fmMultiSelect) As Long
    GetRowFrameWidth = GetTotalColumnWidthNew(ColumnCount, columnDimensions) _
                + m_settings.GetSelectWidth(multiSelect) _
                + m_settings.getRowFramePaddingLeft() _
                 + m_settings.getRowFramePaddingRight()
End Function

Public Function GetScrollHeight(recordsPerPage As Long) As Long
   GetScrollHeight = GetHeightAllRows(recordsPerPage) + m_settings.getScrollHeightPadding()
End Function




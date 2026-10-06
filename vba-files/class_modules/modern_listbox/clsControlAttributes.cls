VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Used to retrieve the attributes for the controls that are part of the Listbox

Private m_controlNameHelpers As New clsControlNameHelpers
Private m_controlPositionCalcs As clsControlPositionCalcs
Private m_settings As clsListBoxSettings

Private m_InitializedCorrectly As Boolean

Private Sub Class_Initialize()
    m_InitializedCorrectly = False
End Sub

Public Property Let InitializedCorrectly(newInitializedCorrectly As Boolean)
    m_InitializedCorrectly = True
End Property

'Public Property Set ParentFrame(argParentFrame As frame)
'    'Set m_controlPositionCalcs = CreateControlPositionCalcs(ParentFrame, m_settings)
'End Property

Public Property Set ControlPositionCalcs(argCalcs As clsControlPositionCalcs)
    Set m_controlPositionCalcs = argCalcs
End Property
Public Property Set settings(argSettings As clsListBoxSettings)
    Set m_settings = argSettings
End Property

Private Function GetControlPositionCalcs() As clsControlPositionCalcs
    If m_InitializedCorrectly = False Then
        Err.Raise errObjectNotSet, "clsControlAttributes.GetControlPositionCalcs" _
    , "clsControlAttributes should be instantiated using CreateControlAttributes() so that the parent frame can be set"
    End If
    Set GetControlPositionCalcs = m_controlPositionCalcs
End Function

Public Function GetCheckboxLabelAttributes(checkFrame As MSForms.frame, row As Long) As labelAttributes
    
    With GetCheckboxLabelAttributes
        .name = m_controlNameHelpers.BuildCheckboxLabelName(row)
        .fontName = "Wingdings 2" 'myicons
        .FontBold = True
        .fontSize = m_settings.dataFontSize() + 4
        .forecolor = rgb(127, 181, 216)
        .caption = "" 'ChrW(&HE03B)
        .Height = checkFrame.Height
        .tag = checkFrame.tag
        .width = checkFrame.width
    End With
    
End Function

Public Function GetCheckboxLabelHeaderAttributes(checkFrame As MSForms.frame) As labelAttributes
    
    With GetCheckboxLabelHeaderAttributes
        .name = m_controlNameHelpers.BuildCheckboxLabelNameHeader
        '.fontName = "Wingdings 2"
        .fontName = "Wingdings 2" 'myicons
        .FontBold = True
        .fontSize = 16
        .forecolor = rgb(127, 181, 216)
        .caption = "" 'ChrW(&HE03B)
        .Height = checkFrame.Height
        .tag = checkFrame.tag
        .width = checkFrame.width
    End With
    
End Function

Public Function GetLabelDataAttributesFixed() As labelAttributes

    With GetLabelDataAttributesFixed

        ' fixed
        .forecolor = m_settings.dataFontColor()
        .fontName = m_settings.dataFontName()
        .fontSize = m_settings.dataFontSize()
        .fontUnderline = m_settings.getDataFontUnderline()
        .Height = m_settings.getDataLabelHeight() + 4 'height of label is 4 more than height of row to allow for padding
        .textAlign = m_settings.getDataLabelTextAlign()
        .top = m_settings.getDataLabelTop()
        
    End With

End Function
 
Public Function GetLabelDataAttributesVariable(ByRef attributes As labelAttributes _
                                , row As Long _
                                , column As Long _
                                , columnDimensions() As ColumnDimension _
                                , multiSelect As fmMultiSelect) As labelAttributes

    With attributes
        ' Variable
        .backcolor = m_settings.getRowBackgroundColor

        .left = GetControlPositionCalcs().GetColumnLeft(column, columnDimensions, multiSelect)
        
        .name = m_controlNameHelpers.BuildRowLabelName(row, column)
        .tag = row
        .width = GetControlPositionCalcs().GetColumnWidth(column, columnDimensions)

    End With
    GetLabelDataAttributesVariable = attributes
End Function

Public Function GetLabelHeaderAttributes(column As Long, columnDimensions() As ColumnDimension) As labelAttributes
    With GetLabelHeaderAttributes
        ' Variable
        .backcolor = m_settings.columnHeaderBackgroundColor
        .left = 0
        .name = m_controlNameHelpers.BuildHeaderLabelName(column)
        .tag = 0
        .width = GetControlPositionCalcs().GetColumnWidth(column, columnDimensions)
        
        ' fixed
        .forecolor = m_settings.columnHeaderFontColor()
        .fontName = m_settings.columnHeaderFontName()
        .fontSize = m_settings.columnHeaderFontSize()
        .Height = m_settings.getDataLabelHeight() + 4 'height of label is 4 more than height of row to allow for padding
        .top = m_settings.getColumnHeaderLabelTop()
        .textAlign = m_settings.getDataLabelTextAlign()

    End With

End Function

Public Function GetLabelHeaderSortAttributes(labelFrameWidth As Long, column As Long) As labelAttributes
    
    With GetLabelHeaderSortAttributes
        .backcolor = m_settings.columnHeaderBackgroundColor
        .caption = ChrW(&H0055)
        .forecolor = m_settings.columnHeaderFontColor()
        .fontName = m_settings.getLabelSortSymbolFontName()
        .fontSize = m_settings.getLabelSortSymbolFontSize()
        .Height = m_settings.getDataLabelHeight() + 4
        .left = GetControlPositionCalcs().GetSymbolLabelLeft(labelFrameWidth)
        .name = m_controlNameHelpers.BuildLabelSymbolName(column)
        .tag = 0
        .textAlign = fmTextAlignLeft
        .top = m_settings.getColumnHeaderLabelTop() + 4
        .width = m_settings.GetLabelSortSymbolWidth()
    End With

End Function

Public Function GetLabelPaginationNumberAttributes(ID As Long, number As Long) As labelAttributes

    With GetLabelPaginationNumberAttributes
        ' Variable
        .backcolor = m_settings.getRowBackgroundColor
        '.caption = CStr(number)
        .left = GetControlPositionCalcs().GetPaginationLabelNumberLeft(number)
        .name = m_controlNameHelpers.BuildPaginationNumberLabelName(ID, number)
        '.tag = number
        
        
        ' fixed
        .forecolor = m_settings.getPaginationLabelFontColor()
        .fontName = m_settings.getPaginationLabelFontName()
        .fontUnderline = m_settings.getPaginationLabelFontUnderline()
        .fontSize = m_settings.getPaginationLabelFontSize()
        .Height = m_settings.getPaginationLabelHeight()
        .textAlign = m_settings.getPaginationLabelTextAlign()
        
        '.top = GetControlPositionCalcs().GetPaginationTop(GetFrameDimensions().top, GetFrameDimensions().height)
        .top = GetControlPositionCalcs().GetPaginationTop()
        .width = m_settings.getPaginationLabelWidth()
        
    End With

End Function

Public Function GetLabelPaginationArrowAttributes(ID As Long, number As Long) As labelAttributes

    With GetLabelPaginationArrowAttributes
        ' Variable
        .backcolor = m_settings.getRowBackgroundColor
        .caption = IIf(number = 1, ">", "<")
       ' .left = GetControlPositionCalcs().GetPaginationLabelArrowLeft(number, GetFrameDimensions().width)
        .left = GetControlPositionCalcs().GetPaginationLabelArrowLeft(number)
        .name = m_controlNameHelpers.BuildPaginationArrowLabelName(ID, number)
        .tag = IIf(number = 1, "right", "left")
        
        ' fixed
        .forecolor = m_settings.getPaginationLabelFontColor()
        .fontName = m_settings.getPaginationLabelFontName()
        .fontSize = m_settings.getPaginationLabelFontSize()
        .Height = m_settings.getPaginationLabelHeight()
        .textAlign = m_settings.getPaginationLabelTextAlign()
        .top = GetControlPositionCalcs().GetPaginationTop()
        .width = m_settings.getPaginationLabelWidth()
        
    End With

End Function

Public Function GetFrameHeaderAttributes(column As Long _
                            , columnDimensions() As ColumnDimension _
                            , multiSelect As fmMultiSelect) As frameHeaderAttributes

    With GetFrameHeaderAttributes
        .name = m_controlNameHelpers.BuildHeaderLabelFrameName(column)
        .caption = ""
        .Height = m_settings.getColumnHeaderRowHeight()
        .left = GetControlPositionCalcs().GetColumnLeft(column, columnDimensions, multiSelect)
        .SpecialEffect = fmSpecialEffectFlat
        .top = 0
        .width = GetControlPositionCalcs().GetColumnWidth(column, columnDimensions)

#If ShowLabels = 1 Then
    .SpecialEffect = fmSpecialEffectEtched ' set to see frame
#End If
    End With
End Function


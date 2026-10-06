Attribute VB_Name = "ClassFactories"
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM

Public Function CreateControlPositionCalcs(parentFrame As clsParentFrameManager _
                                        , settings As clsListBoxSettings) As clsControlPositionCalcs

    Debug.Assert Not settings Is Nothing
    
    Set CreateControlPositionCalcs = New clsControlPositionCalcs
    With CreateControlPositionCalcs
        Set .parentFrame = parentFrame
        Set .settings = settings
        .InitializedCorrectly = True
    End With
End Function

Public Function CreateControlAttributes(settings As clsListBoxSettings _
                                        , calcs As clsControlPositionCalcs) As clsControlAttributes
    Debug.Assert Not settings Is Nothing
    Debug.Assert Not calcs Is Nothing
    
    Set CreateControlAttributes = New clsControlAttributes
    With CreateControlAttributes
        'Set .ParentFrame = ParentFrame
        Set .ControlPositionCalcs = calcs
        Set .settings = settings
        .InitializedCorrectly = True
    End With
    
End Function

Public Function CreateHoverHeader(parentFrame As clsParentFrameManager _
                                            , settings As clsListBoxSettings) As clsHoverHeader
    Debug.Assert Not settings Is Nothing

    Set CreateHoverHeader = New clsHoverHeader
    With CreateHoverHeader
        Set .parentFrame = parentFrame
        Set .settings = settings
        .InitializedCorrectly = True
    End With
    
End Function

Public Function CreateHoverRow(parentFrame As clsParentFrameManager, settings As clsListBoxSettings) As clsHoverRow
    Debug.Assert Not settings Is Nothing

    Set CreateHoverRow = New clsHoverRow
    With CreateHoverRow
        Set .parentFrame = parentFrame
        Set .settings = settings
        .InitializedCorrectly = True
    End With
    
End Function

Public Function CreateHighlightRow(parentFrame As clsParentFrameManager _
                                            , settings As clsListBoxSettings) As clsHighlightRow
    Debug.Assert Not settings Is Nothing

    Set CreateHighlightRow = New clsHighlightRow
    With CreateHighlightRow
        Set .parentFrame = parentFrame
        Set .settings = settings
        .InitializedCorrectly = True
    End With
    
End Function



Attribute VB_Name = "mlbTypes"
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Custom types for the Listbox

Public Type visibleRecordRange
    Start As Long
    end As Long
End Type

Public Type recordDetails
    realRowPosition As Long
    ID As Long
End Type

Public Type checkboxAttributes
    caption As String
    Height As Long
    left As Long
    name As String
    tag As String
    top As Long
    width As Long
End Type

Public Type controlPosition
    left As Long
    width As Long
    top As Long
    Height As Long
End Type

Public Type labelAttributes
    backcolor As Long
    caption As String
    forecolor As Long
    FontBold As Boolean
    fontName As String
    fontSize As Long
    fontUnderline As XlUnderlineStyle
    Height As Long
    left As Long
    name As String
    tag As String
    textAlign As fmTextAlign
    top As Long
    width As Long
End Type

Public Type frameHeaderAttributes
    caption As String
    Height As Long
    left As Long
    name As String
    top As Long
    width As Long
    SpecialEffect As Long
End Type

Public Type ColumnDimension
    left As Long
    width As Long
End Type

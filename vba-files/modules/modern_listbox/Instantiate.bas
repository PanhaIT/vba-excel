Attribute VB_Name = "Instantiate"
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM

' Used to uniquely identity each modern listbox that is created
Private IDTracker As Long

' Needed to instantiate a class from a different VBA project
Public Function Instantiate_clsModernListbox() As clsModernListbox

    Set Instantiate_clsModernListbox = New clsModernListbox
    IDTracker = IDTracker + 1
    Instantiate_clsModernListbox.ID = IDTracker

End Function

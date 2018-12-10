Attribute VB_Name = "CSM_Control_Toolbar"
Option Explicit

Public Function GetTotalWidth(ByRef Toolbar As MSComctlLib.Toolbar) As Single
    Dim Button As MSComctlLib.Button
    Dim TotalWidth As Single
    
    TotalWidth = 0
    For Each Button In Toolbar.Buttons
        TotalWidth = TotalWidth + Button.Width
    Next Button
    
    GetTotalWidth = TotalWidth
End Function

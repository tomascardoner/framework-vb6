Attribute VB_Name = "CSM_Internet"
Option Explicit

'=========================================================================
'INTERNET ADDRESS
Private Const SW_SHOW = 5       'Displays Window in its current size and position
Private Const SW_SHOWNORMAL = 1 'Restores Window if Minimized or Maximized
 
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'=========================================================================

Public Sub OpenAddress(ByVal hwnd As Long, ByVal Address As String)
    ShellExecute hwnd, "open", Address, "", "", SW_SHOWNORMAL
End Sub

Attribute VB_Name = "CSM_Menu"
Option Explicit

'This project needs a form with a menu with at least one submenu
'It also needs a picturebox, Picture1, that contains a small b/w bitmap
Private Const MF_BYPOSITION = &H400&

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Function SetIcon(ByVal FormhWnd As Long)
    Dim hMenu As Long
    Dim hSubMenu As Long
    
    'Get the handle of the menu
    hMenu = GetMenu(FormhWnd)
    If hMenu <> 0 Then
        'Get the first submenu
        hSubMenu = GetSubMenu(hMenu, 2)
        If hSubMenu <> 0 Then
            'Set the menu bitmap
            SetMenuItemBitmaps hSubMenu, 2, MF_BYPOSITION, frmMDI.ilsMenu.ListImages(1).ExtractIcon, frmMDI.ilsMenu.ListImages(1).ExtractIcon
        End If
    End If
End Function

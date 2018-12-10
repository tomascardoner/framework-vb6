Attribute VB_Name = "CSM_Window"
Option Explicit

'=========================================================================
'WINDOW VISIBILITY
Public Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'=========================================================================

'=========================================================================
'WINDOW SIZE AND POSITION
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'=========================================================================

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

'Send a Windows Message to a Window Class
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Const EM_REPLACESEL = &HC2
Public Const EM_SETSEL = &HB1
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const VK_SPACE = &H20
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Private Const MF_GRAYED = &H1&
Private Const MF_BYCOMMAND = &H0&
Private Const SC_CLOSE = &HF060&

'SetMenuItemInfo fMask constants.
Private Const MIIM_STATE     As Long = &H1&
Private Const MIIM_ID        As Long = &H2&

'SetMenuItemInfo fState constants.
Private Const MFS_GRAYED     As Long = &H3&
'Private Const MFS_CHECKED    As Long = &H8&

Private Const WM_NCACTIVATE  As Long = &H86

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

'=========================================================================
'SYSTEM TRAY ICON
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = 0
Private Const NIM_MODIFY = 1
Private Const NIM_DELETE = 2
Private Const NIF_MESSAGE = 1
Private Const NIF_ICON = 2
Private Const NIF_TIP = 4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Private mNID As NOTIFYICONDATA

Private Declare Function Shell_NotifyIconA Lib "shell32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer
'=========================================================================


Public Sub CloseButtonState(ByVal hwnd As Long, ByVal Enable As Boolean)
    Dim wFlags As Long
    Dim hMenu As Long
    Dim Result As Long
    
    hMenu = GetSystemMenu(hwnd, 0)
    If Enable Then
        wFlags = MF_BYCOMMAND And Not MF_GRAYED
    Else
        wFlags = MF_BYCOMMAND Or MF_GRAYED
    End If
    
    Result = EnableMenuItem(hMenu, SC_CLOSE, wFlags)
End Sub

Public Sub CloseMenuAndButtonState(ByVal hwnd As Long, ByVal Enable As Boolean)
    Dim hMenu As Long
    Dim MII As MENUITEMINFO
    Dim Result As Long

    hMenu = GetSystemMenu(hwnd, 0)
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    MII.wID = SC_CLOSE
    Result = GetMenuItemInfo(hMenu, MII.wID, False, MII)
    
    Result = SetId(hMenu, MII, 1)
    If Result = 0 Then
        Exit Sub
    End If
    
    If Enable Then
        MII.fState = MII.fState And Not MFS_GRAYED
    Else
        MII.fState = MII.fState Or MFS_GRAYED
    End If

    MII.fMask = MIIM_STATE
    Result = SetMenuItemInfo(hMenu, MII.wID, False, MII)

    If Result = 0 Then
        Result = SetId(hMenu, MII, 2)
    End If

    Result = SendMessageByNum(hwnd, WM_NCACTIVATE, True, 0)
End Sub

Private Function SetId(ByVal hMenu As Long, ByRef MII As MENUITEMINFO, ByVal Action As Long) As Long
    Dim MenuID As Long
    Dim ret As Long
    
    Const xSC_CLOSE  As Long = -10
    
    MenuID = MII.wID
    If MII.fState = (MII.fState Or MFS_GRAYED) Then
        If Action = 1 Then
            MII.wID = SC_CLOSE
        Else
            MII.wID = xSC_CLOSE
        End If
    Else
        If Action = 1 Then
            MII.wID = xSC_CLOSE
        Else
            MII.wID = SC_CLOSE
        End If
    End If
    
    MII.fMask = MIIM_ID
    ret = SetMenuItemInfo(hMenu, MenuID, False, MII)
    If ret = 0 Then
        MII.wID = MenuID
    End If
    SetId = ret
End Function

Public Sub IconInTraybar_Add(ByVal CallBackhWnd As Long, ByVal Icon As Long, ByVal Tooltip As String)
    Dim ReturnedValue As Long
    
    With mNID
        .cbSize = Len(mNID)
        .hwnd = CallBackhWnd
        .uID = vbNull
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Icon
        .szTip = Tooltip & vbNullChar
    End With
    
    ReturnedValue = Shell_NotifyIconA(NIM_ADD, mNID)
End Sub

Public Sub IconInTraybar_Delete()
    Dim ReturnedValue As Long
    
    ReturnedValue = Shell_NotifyIconA(NIM_DELETE, mNID)
End Sub

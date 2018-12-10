VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_SystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_ARRANGE = 56   'Flags specifying how the system arranged minimized windows. For more information about minimized windows, see the following Remarks section.
Private Const SM_CLEANBOOT = 67 'Value that specifies how the system was started: 0 Normal boot / 1 Fail-safe boot / 2 Fail-safe with network boot Fail-safe boot (also called SafeBoot, Safe Mode, or Clean Boot) bypasses the user's startup files.
Private Const SM_CMETRICS = 44 'Number of system metrics
Private Const SM_CMONITORS = 80 'Number of display monitors on the desktop. See Remarks for more information. Windows NT and Windows 95:  This value is not supported.
Private Const SM_CMOUSEBUTTONS = 43 'Number of buttons on mouse, or zero if no mouse is installed.
Private Const SM_CXBORDER = 5   'Width of a window border, in pixels. This is equivalent to the SM_CXEDGE value for windows with the 3-D look.
Private Const SM_CXCURSOR = 13  'Width of a cursor, in pixels. The system cannot create cursors of other sizes.
Private Const SM_CXDLGFRAME = 7 'Same as SM_CXFIXEDFRAME.
Private Const SM_CXDOUBLECLK = 36   'Width of the rectangle around the location of a first click in a double-click sequence, in pixels. The second click must occur within the rectangle defined by SM_CXDOUBLECLK and SM_CYDOUBLECLK for the system to consider the two clicks a double-click. (The two clicks must also occur within a specified time.) To set the width of the double-click rectangle, call SystemParametersInfo with SPI_SETDOUBLECLKWIDTH.
Private Const SM_CXDRAG = 68    'Width of a rectangle centered on a drag point to allow for limited movement of the mouse pointer before a drag operation begins, in pixels. It allows the user to click and release the mouse button easily without unintentionally starting a drag operation.
Private Const SM_CXEDGE = 45    'Width of a 3-D border, in pixels. This is the 3-D counterpart of SM_CXBORDER.
Private Const SM_CXFIXEDFRAME = 7   'Thickness of the frame around the perimeter of a window that has a caption but is not sizable, in pixels. SM_CXFIXEDFRAME is the height of the horizontal border and SM_CYFIXEDFRAME is the width of the vertical border. Same as SM_CXDLGFRAME.
Private Const SM_CXFOCUSBORDER = 83 'Width of the left and right edges of the focus rectangle drawn by DrawFocusRect. This value is in pixels. Windows 2000/NT and Windows Me/98/95:  This value is not supported.
Private Const SM_CXFRAME = 32   'Same as SM_CXSIZEFRAME.
Private Const SM_CXFULLSCREEN = 16  'Width of the client area for a full-screen window on the primary display monitor, in pixels. To get the coordinates of the portion of the screen not obscured by the system taskbar or by application desktop toolbars, call the SystemParametersInfo function with the SPI_GETWORKAREA value.
Private Const SM_CXHSCROLL = 21 'Width of the arrow bitmap on a horizontal scroll bar, in pixels.
Private Const SM_CXHTHUMB = 10  'Width of the thumb box in a horizontal scroll bar, in pixels.
Private Const SM_CXICON = 11    'Default width of an icon, in pixels. The LoadIcon function can load only icons with the dimensions specified by SM_CXICON and SM_CYICON.
Private Const SM_CXICONSPACING = 38 'Width of a grid cell for items in large icon view, in pixels. Each item fits into a rectangle of size SM_CXICONSPACING by SM_CYICONSPACING when arranged. This value is always greater than or equal to SM_CXICON.
Private Const SM_CXMAXIMIZED = 61   'Default width, in pixels, of a maximized top-level window on the primary display monitor.
Private Const SM_CXMAXTRACK = 59    'Default maximum width of a window that has a caption and sizing borders, in pixels. This metric refers to the entire desktop. The user cannot drag the window frame to a size larger than these dimensions. A window can override this value by processing the WM_GETMINMAXINFO message.
Private Const SM_CXMENUCHECK = 71   'Width of the default menu check-mark bitmap, in pixels.
Private Const SM_CXMENUSIZE = 54    'Width of menu bar buttons, such as the child window close button used in the multiple document interface, in pixels.
Private Const SM_CXMIN = 28 'Minimum width of a window, in pixels.
Private Const SM_CXMINIMIZED = 57   'Width of a minimized window, in pixels.
Private Const SM_CXMINSPACING = 47  'Width of a grid cell for a minimized window, in pixels. Each minimized window fits into a rectangle this size when arranged. This value is always greater than or equal to SM_CXMINIMIZED.
Private Const SM_CXMINTRACK = 34    'Minimum tracking width of a window, in pixels. The user cannot drag the window frame to a size smaller than these dimensions. A window can override this value by processing the WM_GETMINMAXINFO message.
Private Const SM_CXSCREEN = 0   'Width of the screen of the primary display monitor, in pixels. This is the same values obtained by calling GetDeviceCaps as follows: GetDeviceCaps( hdcPrimaryMonitor, HORZRES).
Private Const SM_CXSIZE = 30    'Width of a button in a window's caption or title bar, in pixels.
Private Const SM_CXSIZEFRAME = 32   'Thickness of the sizing border around the perimeter of a window that can be resized, in pixels. SM_CXSIZEFRAME is the width of the horizontal border, and SM_CYSIZEFRAME is the height of the vertical border. Same as SM_CXFRAME.
Private Const SM_CXSMICON = 49  'Recommended width of a small icon, in pixels. Small icons typically appear in window captions and in small icon view.
Private Const SM_CXSMSIZE = 52  'Width of small caption buttons, in pixels.
Private Const SM_CXVIRTUALSCREEN = 76   'Width of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors. The SM_XVIRTUALSCREEN metric is the coordinates of the top of the virtual screen. Windows NT and Windows 95:  This value is not supported.
Private Const SM_CXVSCROLL = 20 'Width of a vertical scroll bar, in pixels.
Private Const SM_CYBORDER = 6   'Height of a window border, in pixels. This is equivalent to the SM_CYEDGE value for windows with the 3-D look.
Private Const SM_CYCAPTION = 4  'Height of a caption area, in pixels.
Private Const SM_CYCURSOR = 14  'Height of a cursor, in pixels. The system cannot create cursors of other sizes.
Private Const SM_CYDLGFRAME = 8 'Same as SM_CYFIXEDFRAME.
Private Const SM_CYDOUBLECLK = 37   'Height of the rectangle around the location of a first click in a double-click sequence, in pixels. The second click must occur within the rectangle defined by SM_CXDOUBLECLK and SM_CYDOUBLECLK for the system to consider the two clicks a double-click. (The two clicks must also occur within a specified time.) To set the height of the double-click rectangle, call SystemParametersInfo with SPI_SETDOUBLECLKHEIGHT.
Private Const SM_CYDRAG = 69    'Height of a rectangle centered on a drag point to allow for limited movement of the mouse pointer before a drag operation begins. This value is in pixels. It allows the user to click and release the mouse button easily without unintentionally starting a drag operation.
Private Const SM_CYEDGE = 46    'Height of a 3-D border, in pixels. This is the 3-D counterpart of SM_CYBORDER.
Private Const SM_CYFIXEDFRAME = 8   'Thickness of the frame around the perimeter of a window that has a caption but is not sizable, in pixels. SM_CXFIXEDFRAME is the height of the horizontal border and SM_CYFIXEDFRAME is the width of the vertical border. Same as SM_CYDLGFRAME.
Private Const SM_CYFOCUSBORDER = 84 'Height of the top and bottom edges of the focus rectangle drawn by DrawFocusRect. This value is in pixels. Windows 2000/NT and Windows Me/98/95:  This value is not supported.
Private Const SM_CYFRAME = 33   'Same as SM_CYSIZEFRAME.
Private Const SM_CYFULLSCREEN = 17  'Height of the client area for a full-screen window on the primary display monitor, in pixels. To get the coordinates of the portion of the screen not obscured by the system taskbar or by application desktop toolbars, call the SystemParametersInfo function with the SPI_GETWORKAREA value.
Private Const SM_CYHSCROLL = 3  'Height of a horizontal scroll bar, in pixels.
Private Const SM_CYICON = 12    'Default height of an icon, in pixels. The LoadIcon function can load only icons with the dimensions SM_CXICON and SM_CYICON.
Private Const SM_CYICONSPACING = 39 'Height of a grid cell for items in large icon view, in pixels. Each item fits into a rectangle of size SM_CXICONSPACING by SM_CYICONSPACING when arranged. This value is always greater than or equal to SM_CYICON.
Private Const SM_CYKANJIWINDOW = 18 'For double byte character set versions of the system, this is the height of the Kanji window at the bottom of the screen, in pixels.
Private Const SM_CYMAXIMIZED = 62   'Default height, in pixels, of a maximized top-level window on the primary display monitor.
Private Const SM_CYMAXTRACK = 60    'Default maximum height of a window that has a caption and sizing borders, in pixels. This metric refers to the entire desktop. The user cannot drag the window frame to a size larger than these dimensions. A window can override this value by processing the WM_GETMINMAXINFO message.
Private Const SM_CYMENU = 15    'Height of a single-line menu bar, in pixels.
Private Const SM_CYMENUCHECK = 72   'Height of the default menu check-mark bitmap, in pixels.
Private Const SM_CYMENUSIZE = 55    'Height of menu bar buttons, such as the child window close button used in the multiple document interface, in pixels.
Private Const SM_CYMIN = 29 'Minimum height of a window, in pixels.
Private Const SM_CYMINIMIZED = 58   'Height of a minimized window, in pixels.
Private Const SM_CYMINSPACING = 48  'Height of a grid cell for a minimized window, in pixels. Each minimized window fits into a rectangle this size when arranged. This value is always greater than or equal to SM_CYMINIMIZED.
Private Const SM_CYMINTRACK = 35    'Minimum tracking height of a window, in pixels. The user cannot drag the window frame to a size smaller than these dimensions. A window can override this value by processing the WM_GETMINMAXINFO message.
Private Const SM_CYSCREEN = 1   'Height of the screen of the primary display monitor, in pixels. This is the same values obtained by calling GetDeviceCaps as follows: GetDeviceCaps( hdcPrimaryMonitor, VERTRES).
Private Const SM_CYSIZE = 31 'height of title bar bitmaps
Private Const SM_CYSIZEFRAME = 33   'Thickness of the sizing border around the perimeter of a window that can be resized, in pixels. SM_CXSIZEFRAME is the width of the horizontal border, and SM_CYSIZEFRAME is the height of the vertical border. Same as SM_CYFRAME.
Private Const SM_CYSMCAPTION = 51   'Height of a small caption, in pixels.
Private Const SM_CYSMICON = 50  'Recommended height of a small icon, in pixels. Small icons typically appear in window captions and in small icon view.
Private Const SM_CYSMSIZE = 53  'Height of small caption buttons, in pixels.
Private Const SM_CYVIRTUALSCREEN = 77   'Height of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors. The SM_YVIRTUALSCREEN metric is the coordinate of the left side of the virtual screen. Windows NT and Windows 95:  This value is not supported.
Private Const SM_CYVSCROLL = 20 'Height of the arrow bitmap on a vertical scroll bar, in pixels.
Private Const SM_CYVTHUMB = 9   'Height of the thumb box in a vertical scroll bar, in pixels.
Private Const SM_DBCSENABLED = 42   'Nonzero if User32.dll supports DBCS; zero otherwise. Windows Me/98/95:  Nonzero if the double-byte character-set (DBCS) version of User.exe is installed; zero otherwise.
Private Const SM_DEBUG = 22 'Nonzero if the debug version of User.exe is installed; zero otherwise.
Private Const SM_IMMENABLED = 82    'Nonzero if Input Method Manager/Input Method Editor features are enabled; zero otherwise. Windows NT and Windows Me/98/95:  This value is not supported. SM_IMMENABLED indicates whether the system is ready to use a Unicode-based IME on a Unicode application. To ensure that a language-dependent IME works, check SM_DBCSENABLED and the system ANSI code page. Otherwise the ANSI-to-Unicode conversion may not be performed correctly, or some components like fonts or registry setting may not be present.
Private Const SM_MEDIACENTER = 87   'Nonzero if the current operating system is the Windows XP, Media Center Edition, zero if not.
Private Const SM_MENUDROPALIGNMENT = 40 'Nonzero if drop-down menus are right-aligned with the corresponding menu-bar item; zero if the menus are left-aligned.
Private Const SM_MIDEASTENABLED = 74    'Nonzero if the system is enabled for Hebrew and Arabic languages, zero if not.
Private Const SM_MOUSEPRESENT = 19  'Nonzero if a mouse is installed; zero otherwise. This value is rarely zero, because of support for virtual mice and because some systems detect the presence of the port instead of the presence of a mouse.
Private Const SM_MOUSEWHEELPRESENT = 19 'Nonzero if a mouse with a wheel is installed; zero otherwise. Windows NT 3.51 and earlier and Windows 95:  This value is not supported.
Private Const SM_NETWORK = 63   'Least significant bit is set if a network is present; otherwise, it is cleared. The other bits are reserved for future use.
Private Const SM_PENWINDOWS = 41    'Nonzero if the Microsoft Windows for Pen computing extensions are installed; zero otherwise.
Private Const SM_REMOTECONTROL = &H2001 'This system metric is used in a Terminal Services environment. Its value is nonzero if the current session is remotely controlled; zero otherwise. Windows 2000/NT and Windows Me/98/95:  This value is not supported.
Private Const SM_REMOTESESSION = &H1000 'This system metric is used in a Terminal Services environment. If the calling process is associated with a Terminal Services client session, the return value is nonzero. If the calling process is associated with the Terminal Server console session, the return value is zero. The console session is not necessarily the physical console - see WTSGetActiveConsoleSessionId for more information. Windows NT 4.0 SP3 and earlier and Windows Me/98/95:  This value is not supported.
Private Const SM_SAMEDISPLAYFORMAT = 81 'Nonzero if all the display monitors have the same color format, zero otherwise. Note that two displays can have the same bit depth, but different color formats. For example, the red, green, and blue pixels can be encoded with different numbers of bits, or those bits can be located in different places in a pixel's color value. Windows NT and Windows 95:  This value is not supported.
Private Const SM_SECURE = 44    'Nonzero if security is present; zero otherwise.
Private Const SM_SHOWSOUNDS = 70    'Nonzero if the user requires an application to present information visually in situations where it would otherwise present the information only in audible form; zero otherwise.
Private Const SM_SHUTTINGDOWN = &H2000  'Nonzero if the current session is shutting down; zero otherwise. Windows 2000/NT and Windows Me/98/95:  This value is not supported.
Private Const SM_SLOWMACHINE = 73   'Nonzero if the computer has a low-end (slow) processor; zero otherwise.
Private Const SM_SWAPBUTTON = 23    'Nonzero if the meanings of the left and right mouse buttons are swapped; zero otherwise.
Private Const SM_TABLETPC = 86  'Nonzero if the current operating system is the Windows XP Tablet PC edition, zero if not.
Private Const SM_XVIRTUALSCREEN = 76    'Coordinates for the left side of the virtual screen. The virtual screen is the bounding rectangle of all display monitors. The SM_CXVIRTUALSCREEN metric is the width of the virtual screen. Windows NT and Windows 95:  This value is not supported.
Private Const SM_YVIRTUALSCREEN = 77    'Coordinates for the top of the virtual screen. The virtual screen is the bounding rectangle of all display monitors. The SM_CYVIRTUALSCREEN metric is the height of the virtual screen. Windows NT and Windows 95:  This value is not supported.

'/////////////////////////////////////////////////////////
'SCREEN

Public Property Get ScreenSizeWidth() As Long
    ScreenSizeWidth = GetSystemMetrics(SM_CXSCREEN)
End Property

Public Property Get ScreenSizeHeight() As Long
    ScreenSizeHeight = GetSystemMetrics(SM_CYSCREEN)
End Property

Public Property Get ScreenVirtualWidth() As Long
    ScreenVirtualWidth = GetSystemMetrics(SM_XVIRTUALSCREEN)
End Property

Public Property Get ScreenVirtualHeight() As Long
    ScreenVirtualHeight = GetSystemMetrics(SM_YVIRTUALSCREEN)
End Property

'/////////////////////////////////////////////////////////
'HORIZONTAL SCROLL BAR

Public Property Get ScrollBarHorizontalArrowSizeWidth() As Long
    ScrollBarHorizontalArrowSizeWidth = GetSystemMetrics(SM_CXHSCROLL)
End Property

Public Property Get ScrollBarHorizontalArrowSizeHeight() As Long
    ScrollBarHorizontalArrowSizeHeight = GetSystemMetrics(SM_CYHSCROLL)
End Property

Public Property Get ScrollBarHorizontalScrollBoxWidth() As Long
    ScrollBarHorizontalScrollBoxWidth = GetSystemMetrics(SM_CXHTHUMB)
End Property

'/////////////////////////////////////////////////////////
'VERTICAL SCROLL BAR

Public Property Get ScrollBarVerticalArrowSizeWidth() As Long
    ScrollBarVerticalArrowSizeWidth = GetSystemMetrics(SM_CXVSCROLL)
End Property

Public Property Get ScrollBarVerticalArrowSizeHeight() As Long
    ScrollBarVerticalArrowSizeHeight = GetSystemMetrics(SM_CYVSCROLL)
End Property

Public Property Get ScrollBarVerticalScrollBoxHeight() As Long
    ScrollBarVerticalScrollBoxHeight = GetSystemMetrics(SM_CYVTHUMB)
End Property

'/////////////////////////////////////////////////////////
'WINDOW

Public Property Get WindowCaptionHeight() As Long
    WindowCaptionHeight = GetSystemMetrics(SM_CYCAPTION)
End Property

Public Property Get WindowSmallCaptionHeight() As Long
    WindowSmallCaptionHeight = GetSystemMetrics(SM_CYSMCAPTION)
End Property

Public Property Get WindowBorderFixedWidth() As Long
    WindowBorderFixedWidth = GetSystemMetrics(SM_CXBORDER)
End Property

Public Property Get WindowBorderFixedHeight() As Long
    WindowBorderFixedHeight = GetSystemMetrics(SM_CYBORDER)
End Property

Public Property Get WindowDialogBorderWitdh() As Long
    WindowDialogBorderWitdh = GetSystemMetrics(SM_CXDLGFRAME)
End Property

Public Property Get WindowDialogBorderHeight() As Long
    WindowDialogBorderHeight = GetSystemMetrics(SM_CXDLGFRAME)
End Property

Public Property Get WindowMaximizedClientAreaWidth() As Long
    WindowMaximizedClientAreaWidth = GetSystemMetrics(SM_CXFULLSCREEN)
End Property

Public Property Get WindowMaximizedClientAreaHeight() As Long
    WindowMaximizedClientAreaHeight = GetSystemMetrics(SM_CYFULLSCREEN)
End Property

Public Property Get WindowKanjiHeight() As Long
    WindowKanjiHeight = GetSystemMetrics(SM_CYKANJIWINDOW)
End Property

Public Property Get WindowMinimumWidth() As Long
    WindowMinimumWidth = GetSystemMetrics(SM_CXMIN)
End Property

Public Property Get WindowMinimumHeight() As Long
    WindowMinimumHeight = GetSystemMetrics(SM_CYMIN)
End Property

Public Property Get WindowTitleBitmapWidth() As Long
    WindowTitleBitmapWidth = GetSystemMetrics(SM_CXSIZE)
End Property

Public Property Get WindowTitleBitmapHeight() As Long
    WindowTitleBitmapHeight = GetSystemMetrics(SM_CYSIZE)
End Property

Public Property Get WindowMaximumTrackingWidth() As Long
    WindowMaximumTrackingWidth = GetSystemMetrics(SM_CXMAXTRACK)
End Property

Public Property Get WindowMaximumTrackingHeight() As Long
    WindowMaximumTrackingHeight = GetSystemMetrics(SM_CYMAXTRACK)
End Property

Public Property Get WindowMinimumTrackingWidth() As Long
    WindowMinimumTrackingWidth = GetSystemMetrics(SM_CXMINTRACK)
End Property

Public Property Get WindowMinimumTrackingHeight() As Long
    WindowMinimumTrackingHeight = GetSystemMetrics(SM_CYMINTRACK)
End Property

Public Property Get WindowMaximizedWidth() As Long
    WindowMaximizedWidth = GetSystemMetrics(SM_CXMAXIMIZED)
End Property

Public Property Get WindowMaximizedHeight() As Long
    WindowMaximizedHeight = GetSystemMetrics(SM_CYMAXIMIZED)
End Property

Public Property Get WindowMinimizedWidth() As Long
    WindowMinimizedWidth = GetSystemMetrics(SM_CXMINIMIZED)
End Property

Public Property Get WindowMinimizedHeight() As Long
    WindowMinimizedHeight = GetSystemMetrics(SM_CYMINIMIZED)
End Property

'/////////////////////////////////////////////////////////
'ICON

Public Property Get IconWidth() As Long
    IconWidth = GetSystemMetrics(SM_CXICON)
End Property

Public Property Get IconHeight() As Long
    IconHeight = GetSystemMetrics(SM_CYICON)
End Property

Public Property Get IconSpacingWidth() As Long
    IconSpacingWidth = GetSystemMetrics(SM_CXICONSPACING)
End Property

Public Property Get IconSpacingHeight() As Long
    IconSpacingHeight = GetSystemMetrics(SM_CYICONSPACING)
End Property

'/////////////////////////////////////////////////////////
'CURSOR

Public Property Get CursorWidth() As Long
    CursorWidth = GetSystemMetrics(SM_CXCURSOR)
End Property

Public Property Get CursorHeight() As Long
    CursorHeight = GetSystemMetrics(SM_CYCURSOR)
End Property

'/////////////////////////////////////////////////////////
'MENU

Public Property Get MenuHeight() As Long
    MenuHeight = GetSystemMetrics(SM_CYMENU)
End Property

Public Property Get MenuDropAlignmentLeft() As Boolean
    MenuDropAlignmentLeft = Not GetSystemMetrics(SM_MENUDROPALIGNMENT)
End Property

Public Property Get MenuCheckBitmapWidth() As Long
    MenuCheckBitmapWidth = GetSystemMetrics(SM_CXMENUCHECK)
End Property

Public Property Get MenuCheckBitmapHeight() As Long
    MenuCheckBitmapHeight = GetSystemMetrics(SM_CYMENUCHECK)
End Property

Public Property Get MenuButtonWidth() As Long
    MenuButtonWidth = GetSystemMetrics(SM_CXMENUSIZE)
End Property

Public Property Get MenuButtonHeight() As Long
    MenuButtonHeight = GetSystemMetrics(SM_CYMENUSIZE)
End Property

'/////////////////////////////////////////////////////////
'MOUSE

Public Property Get MousePresent() As Boolean
    MousePresent = GetSystemMetrics(SM_MOUSEPRESENT)
End Property

Public Property Get MouseButtonsSwapped() As Boolean
    MouseButtonsSwapped = GetSystemMetrics(SM_SWAPBUTTON)
End Property

Public Property Get MouseDoubleClickWidth() As Long
    MouseDoubleClickWidth = GetSystemMetrics(SM_CXDOUBLECLK)
End Property

Public Property Get MouseDoubleClickHeight() As Long
    MouseDoubleClickHeight = GetSystemMetrics(SM_CYDOUBLECLK)
End Property

Public Property Get MouseButtonsCount() As Long
    MouseButtonsCount = GetSystemMetrics(SM_CMOUSEBUTTONS)
End Property

Public Property Get MouseDragAreaWidth() As Long
    MouseDragAreaWidth = GetSystemMetrics(SM_CXDRAG)
End Property

Public Property Get MouseDragAreaHeight() As Long
    MouseDragAreaHeight = GetSystemMetrics(SM_CYDRAG)
End Property

'/////////////////////////////////////////////////////////
'WINDOWS SYSTEM

Public Property Get WindowsSystemDebug() As Boolean
    WindowsSystemDebug = GetSystemMetrics(SM_DEBUG)
End Property

Public Property Get WindowsSystemPenInstalled() As Boolean
    WindowsSystemPenInstalled = GetSystemMetrics(SM_PENWINDOWS)
End Property

Public Property Get WindowsSystemIsTabletPCEdition() As Boolean
    WindowsSystemIsTabletPCEdition = GetSystemMetrics(SM_TABLETPC)
End Property

Public Property Get WindowsSystemDoubleByteCharacterSystemEnabled() As Boolean
    WindowsSystemDoubleByteCharacterSystemEnabled = GetSystemMetrics(SM_DBCSENABLED)
End Property

Public Property Get WindowsSystemMidEastEnabled() As Boolean
    WindowsSystemMidEastEnabled = GetSystemMetrics(SM_MIDEASTENABLED)
End Property

Public Property Get WindowsSystemMetricsCount() As Long
    WindowsSystemMetricsCount = GetSystemMetrics(SM_CMETRICS)
End Property

Public Property Get WindowsSystemSafeMode() As Boolean
    WindowsSystemSafeMode = GetSystemMetrics(SM_CLEANBOOT)
End Property

Public Property Get WindowsSystemShuttingDown() As Boolean
    WindowsSystemShuttingDown = GetSystemMetrics(SM_SHUTTINGDOWN)
End Property

Public Property Get WindowsSystemNetworkPresent() As Boolean
    WindowsSystemNetworkPresent = GetSystemMetrics(SM_NETWORK)
End Property

Public Property Get WindowsSystemSecure() As Boolean
    WindowsSystemSecure = GetSystemMetrics(SM_SECURE)
End Property

Public Property Get WindowsSystemSlowMachine() As Boolean
    WindowsSystemSlowMachine = GetSystemMetrics(SM_SLOWMACHINE)
End Property

Public Property Get WindowsSystemIsRemoteSession() As Boolean
    WindowsSystemIsRemoteSession = GetSystemMetrics(SM_REMOTESESSION)
End Property

Public Property Get WindowsSystemIsRemoteControlled() As Boolean
    WindowsSystemIsRemoteControlled = GetSystemMetrics(SM_REMOTECONTROL)
End Property

Public Property Get WindowsSystemMonitorsCount() As Long
    WindowsSystemMonitorsCount = GetSystemMetrics(SM_CMONITORS)
End Property

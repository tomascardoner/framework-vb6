Attribute VB_Name = "CSM_System"
Option Explicit

'//////////  SYSTEM METRICS  \\\\\\\\\\
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CMETRICS = 44
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CXBORDER = 5
Public Const SM_CXCURSOR = 13
Public Const SM_CXDLGFRAME = 7
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CXEDGE = 45
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CXFRAME = 32
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CXHSCROLL = 21
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CXICONSPACING = 38
Public Const SM_CXMIN = 28
Public Const SM_CXMINTRACK = 34
Public Const SM_CXPADDEDBORDER = 92
Public Const SM_CXSCREEN = 0
Public Const SM_CXSIZE = 30
Public Const SM_CXSIZEFRAME = SM_CXFRAME
Public Const SM_CXVSCROLL = 2
Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4
Public Const SM_CYCURSOR = 14
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CYEDGE = 46
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYFRAME = 33
Public Const SM_CYHSCROLL = 3
Public Const SM_CYICON = 12
Public Const SM_CYICONSPACING = 39
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_CYMENU = 15
Public Const SM_CYMIN = 29
Public Const SM_CYMINTRACK = 35
Public Const SM_CYSCREEN = 1
Public Const SM_CYSIZE = 31
Public Const SM_CYSIZEFRAME = SM_CYFRAME
Public Const SM_CYVSCROLL = 20
Public Const SM_CYVTHUMB = 9
Public Const SM_DBCSENABLED = 42
Public Const SM_DEBUG = 22
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_MOUSEPRESENT = 19
Public Const SM_PENWINDOWS = 41

'//////////  WINDOWS VERSION  \\\\\\\\\\
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long

' Returns the version of Windows that the user is running
Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"

                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista/Server 2008"
                            Case 1
                                GetWindowsVersion = "Windows 7/Server 2008 R2"
                        End Select
                End Select

            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Unable to identify your version of Windows."
    End If
End Function

Public Function IsHost64Bit() As Boolean
    Dim handle As Long
    Dim is64Bit As Boolean

    ' Assume initially that this is not a WOW64 process
    is64Bit = False

    ' Then try to prove that wrong by attempting to load the
    ' IsWow64Process function dynamically
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    ' The function exists, so call it
    If handle <> 0 Then
        IsWow64Process GetCurrentProcess(), is64Bit
    End If

    ' Return the value
    IsHost64Bit = is64Bit
End Function


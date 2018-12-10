VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_SpecialFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Declare Function SHGetFolderPath Lib "shell32" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Private Const SHGFP_TYPE_CURRENT As Long = 0
Private Const SHGFP_TYPE_DEFAULT  As Long = 1

Private Const S_FALSE = &H1
Private Const E_FAIL = &H80004005
Private Const E_INVALIDARG = &H80070057

Private Const CSIDL_ADMINTOOLS As Long = &H30               'Version 5.0. The file system directory that is used to store administrative tools for an individual user. The Microsoft Management Console (MMC) will save customized consoles to this directory, and it will roam with the user.
Private Const CSIDL_ALTSTARTUP As Long = &H1D               'The file system directory that corresponds to the user's nonlocalized Startup program group.
Private Const CSIDL_APPDATA As Long = &H1A                  'Version 4.71. The file system directory that serves as a common repository for application-specific data. A typical path is C:\Documents and Settings\username\Application Data. This CSIDL is supported by the redistributable Shfolder.dll for systems that do not have the Microsoft Internet Explorer 4.0 integrated Shell installed.
Private Const CSIDL_BITBUCKET As Long = &HA                 'The virtual folder containing the objects in the user's Recycle Bin.
Private Const CSIDL_CDBURN_AREA As Long = &H3B              'Version 6.0. The file system directory acting as a staging area for files waiting to be written to CD. A typical path is C:\Documents and Settings\username\Local Settings\Application Data\Microsoft\CD Burning.
Private Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F        'Version 5.0. The file system directory containing administrative tools for all users of the computer.
Private Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E        'The file system directory that corresponds to the nonlocalized Startup program group for all users. Valid only for Microsoft Windows NT systems.
Private Const CSIDL_COMMON_APPDATA As Long = &H23           'Version 5.0. The file system directory containing application data for all users. A typical path is C:\Documents and Settings\All Users\Application Data.
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19  'The file system directory that contains files and folders that appear on the desktop for all users. A typical path is C:\Documents and Settings\All Users\Desktop. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_DOCUMENTS As Long = &H2E         'The file system directory that contains documents that are common to all users. A typical paths is C:\Documents and Settings\All Users\Documents. Valid for Windows NT systems and Microsoft Windows 95 and Windows 98 systems with Shfolder.dll installed.
Private Const CSIDL_COMMON_FAVORITES As Long = &H1F         'The file system directory that serves as a common repository for favorite items common to all users. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_MUSIC As Long = &H35             'Version 6.0. The file system directory that serves as a repository for music files common to all users. A typical path is C:\Documents and Settings\All Users\Documents\My Music.
Private Const CSIDL_COMMON_PICTURES As Long = &H36          'Version 6.0. The file system directory that serves as a repository for image files common to all users. A typical path is C:\Documents and Settings\All Users\Documents\My Pictures.
Private Const CSIDL_COMMON_PROGRAMS As Long = &H17          'The file system directory that contains the directories for the common program groups that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu\Programs. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_STARTMENU As Long = &H16         'The file system directory that contains the programs and folders that appear on the Start menu for all users. A typical path is C:\Documents and Settings\All Users\Start Menu. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_STARTUP As Long = &H18           'The file system directory that contains the programs that appear in the Startup folder for all users. A typical path is C:\Documents and Settings\All Users\Start Menu\Programs\Startup. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_TEMPLATES As Long = &H2D         'The file system directory that contains the templates that are available to all users. A typical path is C:\Documents and Settings\All Users\Templates. Valid only for Windows NT systems.
Private Const CSIDL_COMMON_VIDEO As Long = &H37             'Version 6.0. The file system directory that serves as a repository for video files common to all users. A typical path is C:\Documents and Settings\All Users\Documents\My Videos.
Private Const CSIDL_CONTROLS As Long = &H3                  'The virtual folder containing icons for the Control Panel applications.
Private Const CSIDL_COOKIES As Long = &H21                  'The file system directory that serves as a common repository for Internet cookies. A typical path is C:\Documents and Settings\username\Cookies.
Private Const CSIDL_DESKTOP As Long = &H0                   'The virtual folder representing the Windows desktop, the root of the namespace.
Private Const CSIDL_DESKTOPDIRECTORY As Long = &H10         'The file system directory used to physically store file objects on the desktop (not to be confused with the desktop folder itself). A typical path is C:\Documents and Settings\username\Desktop.
Private Const CSIDL_DRIVES As Long = &H11                   'The virtual folder representing My Computer, containing everything on the local computer: storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
Private Const CSIDL_FAVORITES As Long = &H6                 'The file system directory that serves as a common repository for the user's favorite items. A typical path is C:\Documents and Settings\username\Favorites.
Private Const CSIDL_FLAG_CREATE As Long = &H8000            'Version 5.0. Combine this CSIDL with any of the following CSIDLs to force the creation of the associated folder.
Private Const CSIDL_FONTS As Long = &H14                    'A virtual folder containing fonts. A typical path is C:\Windows\Fonts.
Private Const CSIDL_HISTORY As Long = &H22                  'The file system directory that serves as a common repository for Internet history items.
Private Const CSIDL_INTERNET As Long = &H1                  'A virtual folder representing the Internet.
Private Const CSIDL_INTERNET_CACHE As Long = &H20           'Version 4.72. The file system directory that serves as a common repository for temporary Internet files. A typical path is C:\Documents and Settings\username\Local Settings\Temporary Internet Files.
Private Const CSIDL_LOCAL_APPDATA As Long = &H1C            'Version 5.0. The file system directory that serves as a data repository for local (nonroaming) applications. A typical path is C:\Documents and Settings\username\Local Settings\Application Data.
Private Const CSIDL_MYDOCUMENTS As Long = &HC               'Version 6.0. The virtual folder representing the My Documents desktop item.
Private Const CSIDL_MYMUSIC As Long = &HD                   'The file system directory that serves as a common repository for music files. A typical path is C:\Documents and Settings\User\My Documents\My Music.
Private Const CSIDL_MYPICTURES As Long = &H27               'Version 5.0. The file system directory that serves as a common repository for image files. A typical path is C:\Documents and Settings\username\My Documents\My Pictures.
Private Const CSIDL_MYVIDEO As Long = &HE                   'Version 6.0. The file system directory that serves as a common repository for video files. A typical path is C:\Documents and Settings\username\My Documents\My Videos.
Private Const CSIDL_NETHOOD As Long = &H13                  'A file system directory containing the link objects that may exist in the My Network Places virtual folder. It is not the same as CSIDL_NETWORK, which represents the network namespace root. A typical path is C:\Documents and Settings\username\NetHood.
Private Const CSIDL_NETWORK As Long = &H12                  'A virtual folder representing Network Neighborhood, the root of the network namespace hierarchy.
Private Const CSIDL_PERSONAL As Long = &H5                  'Version 6.0. The virtual folder representing the My Documents desktop item. This is equivalent to CSIDL_MYDOCUMENTS. Previous to Version 6.0. The file system directory used to physically store a user's common repository of documents. A typical path is C:\Documents and Settings\username\My Documents. This should be distinguished from the virtual My Documents folder in the namespace. To access that virtual folder, use SHGetFolderLocation, which returns the ITEMIDLIST for the virtual location, or refer to the technique described in Managing the File System.
Private Const CSIDL_PRINTERS As Long = &H4                  'The virtual folder containing installed printers.
Private Const CSIDL_PRINTHOOD As Long = &H1B                'The file system directory that contains the link objects that can exist in the Printers virtual folder. A typical path is C:\Documents and Settings\username\PrintHood.
Private Const CSIDL_PROFILE As Long = &H28                  'Version 5.0. The user's profile folder. A typical path is C:\Documents and Settings\username. Applications should not create files or folders at this level; they should put their data under the locations referred to by CSIDL_APPDATA or CSIDL_LOCAL_APPDATA.
Private Const CSIDL_PROFILES As Long = &H3E                 'Version 6.0. The file system directory containing user profile folders. A typical path is C:\Documents and Settings.
Private Const CSIDL_PROGRAM_FILES As Long = &H26            'Version 5.0. The Program Files folder. A typical path is C:\Program Files.
Private Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B     'Version 5.0. A folder for components that are shared across applications. A typical path is C:\Program Files\Common. Valid only for Windows NT, Windows 2000, and Windows XP systems. Not valid for Windows Millennium Edition (Windows Me).
Private Const CSIDL_PROGRAMS As Long = &H2                  'The file system directory that contains the user's program groups (which are themselves file system directories). A typical path is C:\Documents and Settings\username\Start Menu\Programs.
Private Const CSIDL_RECENT As Long = &H8                    'The file system directory that contains shortcuts to the user's most recently used documents. A typical path is C:\Documents and Settings\username\My Recent Documents. To create a shortcut in this folder, use SHAddToRecentDocs. In addition to creating the shortcut, this function updates the Shell's list of recent documents and adds the shortcut to the My Recent Documents submenu of the Start menu.
Private Const CSIDL_SENDTO As Long = &H9                    'The file system directory that contains Send To menu items. A typical path is C:\Documents and Settings\username\SendTo.
Private Const CSIDL_STARTMENU As Long = &HB                 'The file system directory containing Start menu items. A typical path is C:\Documents and Settings\username\Start Menu.
Private Const CSIDL_STARTUP As Long = &H7                   'The file system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user logs onto Windows NT or starts Windows 95. A typical path is C:\Documents and Settings\username\Start Menu\Programs\Startup.
Private Const CSIDL_SYSTEM As Long = &H25                   'Version 5.0. The Windows System folder. A typical path is C:\Windows\System32.
Private Const CSIDL_TEMPLATES As Long = &H15                'The file system directory that serves as a common repository for document templates. A typical path is C:\Documents and Settings\username\Templates.
Private Const CSIDL_WINDOWS As Long = &H24                  'Version 5.0. The Windows directory or SYSROOT. This corresponds to the %windir% or %SYSTEMROOT% environment variables. A typical path is C:\Windows.

Private mTemp As String
Private mWindows As String
Private mSystem As String
Private mMyDocuments As String
Private mFavorites As String
Private mFonts As String

Public Property Get Temp() As String
    Temp = mTemp
End Property

Public Property Get Windows() As String
    Windows = mWindows
End Property

Public Property Get System() As String
    System = mSystem
End Property

Public Property Get MyDocuments() As String
    MyDocuments = mMyDocuments
End Property

Public Property Get Favorites() As String
    Favorites = mFavorites
End Property

Public Property Get Fonts() As String
    Fonts = mFonts
End Property

Public Function GetTempPathFileName(Optional Prefix As String = "") As String
    Dim strTempPath As String
    Dim lngResult As Long
    Dim strBuffer As String
    
    strTempPath = mTemp
    If strTempPath = "" Then strTempPath = App.Path
    
    strBuffer = String(255, 0)
    lngResult = GetTempFileName(strTempPath, Prefix, 0, strBuffer)
    If lngResult = 0 Then
        GetTempPathFileName = Prefix + Format(Now, "yyyymmddhhnnss") + ".TMP"
    Else
        Kill strTempPath + Prefix + Hex(lngResult) + ".TMP"
        GetTempPathFileName = strTempPath + Prefix + Hex(lngResult) + ".TMP"
    End If
End Function

Private Sub Class_Initialize()
    Dim lngBufferLength As Long
    Dim strBuffer As String
    Dim lngResult As Long
    
    'TEMP FOLDER
    strBuffer = String(255, 0)
    lngBufferLength = GetTempPath(Len(strBuffer), strBuffer)
    If lngBufferLength = 0 Then
        mTemp = ""
    Else
        mTemp = Left(strBuffer, lngBufferLength)
    End If
    
    'WINDOWS FOLDER
    strBuffer = String(255, 0)
    lngResult = SHGetFolderPath(0, CSIDL_WINDOWS, 0, SHGFP_TYPE_CURRENT, strBuffer)
    If lngResult = 0 Then
        mWindows = CSM_String.RemoveCharsAfterNull(strBuffer)
    End If
    
    'SYSTEM FOLDER
    strBuffer = String(255, 0)
    lngResult = SHGetFolderPath(0, CSIDL_SYSTEM, 0, SHGFP_TYPE_CURRENT, strBuffer)
    If lngResult = 0 Then
        mSystem = CSM_String.RemoveCharsAfterNull(strBuffer)
    End If
    
    'MYDOCUMENTS FOLDER
    strBuffer = String(255, 0)
    lngResult = SHGetFolderPath(0, CSIDL_PERSONAL, 0, SHGFP_TYPE_CURRENT, strBuffer)
    If lngResult = 0 Then
        mMyDocuments = CSM_String.RemoveCharsAfterNull(strBuffer)
    End If
    
    'FAVORITES FOLDER
    strBuffer = String(255, 0)
    lngResult = SHGetFolderPath(0, CSIDL_FAVORITES, 0, SHGFP_TYPE_CURRENT, strBuffer)
    If lngResult = 0 Then
        mFavorites = CSM_String.RemoveCharsAfterNull(strBuffer)
    End If
    
    'FONTS FOLDER
    strBuffer = String(255, 0)
    lngResult = SHGetFolderPath(0, CSIDL_FONTS, 0, SHGFP_TYPE_CURRENT, strBuffer)
    If lngResult = 0 Then
        mFonts = CSM_String.RemoveCharsAfterNull(strBuffer)
    End If
End Sub


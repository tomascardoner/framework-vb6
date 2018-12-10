Attribute VB_Name = "CSM_Drive"
Option Explicit

'Private Const DRIVE_CDROM As Long = 5
'Private Const DRIVE_REMOVABLE As Long = 2
'
'Private Const GENERIC_READ As Long = &H80000000
'Private Const GENERIC_WRITE As Long = &H40000000
'
'Private Const OPEN_EXISTING As Long = 3
'Private Const FILE_DEVICE_FILE_SYSTEM As Long = 9
'Private Const FILE_DEVICE_MASS_STORAGE As Long = &H2D&
'Private Const METHOD_BUFFERED As Long = 0
'Private Const FILE_ANY_ACCESS As Long = 0
'Private Const FILE_READ_ACCESS As Long = 1
'Private Const LOCK_VOLUME As Long = 6
'Private Const DISMOUNT_VOLUME As Long = 8
'Private Const EJECT_MEDIA As Long = &H202
'Private Const MEDIA_REMOVAL As Long = &H201
'Private Const INVALID_HANDLE_VALUE As Long = -1
'
'Private Const LOCK_TIMEOUT As Long = 1000
'Private Const LOCK_RETRIES As Long = 20
'
'Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByRef dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByRef lpOverlapped As Long) As Long
'
'Private boTimeOut As Boolean
'Private mTimer As Timer
'
'Private Function CTL_CODE(lngDevFileSys As Long, lngFunction As Long, lngMethod As Long, lngAccess As Long) As Long
'    CTL_CODE = (lngDevFileSys * (2 ^ 16)) Or (lngAccess * (2 ^ 14)) Or (lngFunction * (2 ^ 2)) Or lngMethod
'End Function
'
'Private Function OpenVolume(strLetter As String, lngVolHandle As Long) As Boolean
'    Dim lngDriveType As Long
'    Dim lngAccessFlags As Long
'    Dim strVolume As String
'
'    lngDriveType = GetDriveType(strLetter)
'    Select Case lngDriveType
'        Case DRIVE_REMOVABLE
'            lngAccessFlags = GENERIC_READ Or GENERIC_WRITE
'        Case DRIVE_CDROM
'            lngAccessFlags = GENERIC_READ
'        Case Else
'            OpenVolume = False
'            Exit Function
'    End Select
'    strVolume = "\\.\" & strLetter
'    lngVolHandle = CreateFile(strVolume, lngAccessFlags, 0, ByVal CLng(0), OPEN_EXISTING, ByVal CLng(0), ByVal CLng(0))
'    If lngVolHandle = INVALID_HANDLE_VALUE Then
'        OpenVolume = False
'        Exit Function
'    End If
'    OpenVolume = True
'End Function
'
'Private Function CloseVolume(lngVolHandle As Long) As Boolean
'    Dim lngReturn As Long
'
'    lngReturn = CloseHandle(lngVolHandle)
'    If lngReturn = 0 Then
'        CloseVolume = False
'    Else
'        CloseVolume = True
'    End If
'End Function
'
'Private Function LockVolume(ByRef lngVolHandle As Long) As Boolean
'    Dim lngBytesReturned As Long
'    Dim intCount As Integer
'    Dim intI As Integer
'    Dim boLocked As Boolean
'    Dim lngFunction As Long
'
'    lngFunction = CTL_CODE(FILE_DEVICE_FILE_SYSTEM, LOCK_VOLUME, METHOD_BUFFERED, FILE_ANY_ACCESS)
'    intCount = LOCK_TIMEOUT / LOCK_RETRIES
'    boLocked = False
'
'    For intI = 0 To LOCK_RETRIES
'        boTimeOut = False
'        mTimer.Interval = intCount
'        mTimer.Enabled = True
'        Do Until boTimeOut = True Or boLocked = True
'            boLocked = DeviceIoControl(lngVolHandle, ByVal lngFunction, CLng(0), 0, CLng(0), 0, lngBytesReturned, ByVal CLng(0))
'            DoEvents
'        Loop
'        If boLocked = True Then
'            LockVolume = True
'            mTimer.Enabled = False
'            Exit Function
'        End If
'    Next intI
'    LockVolume = False
'End Function
'
'Private Function DismountVolume(lngVolHandle As Long) As Boolean
'    Dim lngBytesReturned As Long
'    Dim lngFunction As Long
'
'    lngFunction = CTL_CODE(FILE_DEVICE_FILE_SYSTEM, DISMOUNT_VOLUME, METHOD_BUFFERED, FILE_ANY_ACCESS)
'    DismountVolume = DeviceIoControl(lngVolHandle, ByVal lngFunction, 0, 0, 0, 0, lngBytesReturned, ByVal 0)
'End Function
'
'Private Function PreventRemovalofVolume(lngVolHandle As Long) As Boolean
'    Dim boPreventRemoval As Boolean
'    Dim lngBytesReturned As Long
'    Dim lngFunction As Long
'
'    boPreventRemoval = False
'    lngFunction = CTL_CODE(FILE_DEVICE_MASS_STORAGE, MEDIA_REMOVAL, METHOD_BUFFERED, FILE_READ_ACCESS)
'    PreventRemovalofVolume = DeviceIoControl(lngVolHandle, ByVal lngFunction, boPreventRemoval, Len(boPreventRemoval), 0, 0, lngBytesReturned, ByVal 0)
'End Function
'
'Private Function AutoEjectVolume(lngVolHandle As Long) As Boolean
'    Dim lngFunction As Long
'    Dim lngBytesReturned As Long
'
'    lngFunction = CTL_CODE(FILE_DEVICE_MASS_STORAGE, EJECT_MEDIA, METHOD_BUFFERED, FILE_READ_ACCESS)
'    AutoEjectVolume = DeviceIoControl(lngVolHandle, ByVal lngFunction, 0, 0, 0, 0, lngBytesReturned, ByVal 0)
'End Function
'
'Public Sub Eject(strVol As String)
'    Dim lngVolHand As Long
'    Dim boResult As Boolean
'    Dim boSafe As Boolean
'
'    strVol = strVol & ":"
'    '
'    ' Open and get a Handle for the Volume
'    '
'    boResult = OpenVolume(strVol, lngVolHand)
'    If boResult = False Then
'        ShowErrorMessage "CSM_Drive.Eject", "Error al abrir el volumen. (error: " & Err.LastDllError & ")"
'        Exit Sub
'    End If
'    '
'    ' Lock the Volume
'    '
'    boResult = LockVolume(lngVolHand)
'    If boResult = False Then
'        ShowErrorMessage "CSM_Drive.Eject", "Error al trabar el volumen. (error: " & Err.LastDllError & ")"
'        CloseVolume (lngVolHand)
'        Exit Sub
'    End If
'    '
'    'Dismount the Volume
'    '
'    boResult = DismountVolume(lngVolHand)
'    If boResult = False Then
'        ShowErrorMessage "CSM_Drive.Eject", "Error al desmontar el volumen. (error: " & Err.LastDllError & ")"
'        CloseVolume (lngVolHand)
'        Exit Sub
'    End If
'    '
'    ' Set to allow the Volume to be Removed
'    '
'    boResult = PreventRemovalofVolume(lngVolHand)
'    If boResult = False Then
'        ShowErrorMessage "CSM_Drive.Eject", "Error al permitir la extracción del volumen. (error: " & Err.LastDllError & ")"
'        CloseVolume (lngVolHand)
'        Exit Sub
'    End If
'    boSafe = True
'    '
'    ' Eject the Volume
'    '
'    boResult = AutoEjectVolume(lngVolHand)
'    If boSafe = True Then
'        MsgBox "Ya puede retirar el Volumen " & UCase(strVol)
'    End If
'    '
'    ' Close the Handle
'    '
'    boResult = CloseVolume(lngVolHand)
'    If boResult = False Then
'        ShowErrorMessage "CSM_Drive.Eject", "Error al cerrar el Volumen. (error: " & Err.LastDllError & ")"
'        Exit Sub
'    End If
'End Sub
'
'Private Sub Timer1_Timer()
'    boTimeOut = True
'End Sub

'Public Function ExtractUSBDrive(ByVal DriveLetter As String) As Boolean
'    Dim EjectDevSpec As String
'
'    If Not GetDriveDeviceId(DriveLetter, EjectDevSpec) Then
'    End If
'    ExtractUSBDrive = True
'    Exit Function
'
'End Function
'
'Public Function GetDriveDeviceId(ByVal DriveSpec As String) As String
'    Dim DeviceId As String
'    Dim TruncationPos As Integer
'
'    'Get RAW info
'    DeviceId = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "SYSTEM\\MountedDevices", "\DosDevices\" & DriveSpec & ":", "", csrdtString)
'    'Clean NULL chars
'    DeviceId = Replace(DeviceId, vbNullChar, "")
'
'    If Left(DeviceId, 4) = "\??\" Then
'        DeviceId = Mid(DeviceId, 5)
'        DeviceId = Replace(DeviceId, "#", "\")
'        TruncationPos = InStr(1, DeviceId, "\{")
'        If TruncationPos > 1 Then
'            DeviceId = Left(DeviceId, TruncationPos - 1)
'        End If
'    End If
'    GetDriveDeviceId = DeviceId
'End Function
'
'Public Function GetFriendlyName(ByVal DeviceId As String) As String
'    Dim FriendlyName As String
'
'    FriendlyName = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Enum\" & DeviceId, "FriendlyName", "", csrdtString)
'    If FriendlyName = "" Then
'        FriendlyName = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Enum\" & DeviceId, "DeviceDesc", "", csrdtString)
'    End If
'
'    GetFriendlyName = FriendlyName
'End Function
'

Private Declare Function CM_Get_DevNode_Status Lib "setupapi.dll" (lStatus As Long, lProblem As Long, ByVal hDevice As Long, ByVal dwFlags As Long) As Long
Private Declare Function CM_Get_Parent Lib "setupapi.dll" (hParentDevice As Long, ByVal hDevice As Long, ByVal dwFlags As Long) As Long
Private Declare Function CM_Locate_DevNodeA Lib "setupapi.dll" (hDevice As Long, ByVal lpDeviceName As Long, ByVal dwFlags As Long) As Long
Private Declare Function CM_Request_Device_EjectA Lib "setupapi.dll" (ByVal hDevice As Long, lVetoType As Long, ByVal lpVetoName As Long, ByVal cbVetoName As Long, ByVal dwFlags As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long

' Safely remove USB flash drive
Public Function SafelyRemove(ByVal pstrDrive As String) As Boolean
    Const DN_REMOVABLE = &H4000
    Dim strDeviceInstance As String
    Dim lngDevice As Long
    Dim lngStatus As Long
    Dim lngProblem As Long
    Dim lngVetoType As Long
    Dim strVeto As String * 255
    
    pstrDrive = UCase$(Left$(pstrDrive, 1)) & ":"
    strDeviceInstance = StrConv(GetDeviceInstance(pstrDrive), vbFromUnicode)
    If CM_Locate_DevNodeA(lngDevice, StrPtr(strDeviceInstance), 0) = 0 Then
        If CM_Get_DevNode_Status(lngStatus, lngProblem, lngDevice, 0) = 0 Then
            Do While Not (lngStatus And DN_REMOVABLE) > 0
                If CM_Get_Parent(lngDevice, lngDevice, 0) <> 0 Then Exit Do
                If CM_Get_DevNode_Status(lngStatus, lngProblem, lngDevice, 0) <> 0 Then Exit Do
            Loop
            If (lngStatus And DN_REMOVABLE) > 0 Then SafelyRemove = (CM_Request_Device_EjectA(lngDevice, lngVetoType, StrPtr(strVeto), 255, 0) = 0)
        End If
    End If
End Function

Private Function GetDeviceInstance(pstrDrive As String) As String
    Const HKEY_LOCAL_MACHINE = &H80000002
    Const KEY_QUERY_VALUE = &H1
    Const REG_BINARY = &H3
    Const ERROR_SUCCESS = 0&
    Dim strKey As String
    Dim strValue As String
    Dim lngHandle As Long
    Dim lngType As Long
    Dim strBuffer As String
    Dim lngLen As Long
    Dim bytArray() As Byte
    
    strKey = "SYSTEM\MountedDevices"
    strValue = "\DosDevices\" & pstrDrive
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, strKey, 0&, KEY_QUERY_VALUE, lngHandle) = ERROR_SUCCESS Then
        If RegQueryValueEx(lngHandle, strValue, 0&, lngType, 0&, lngLen) = 234 Then
            If lngType = REG_BINARY Then
                strBuffer = Space$(lngLen)
                If RegQueryValueEx(lngHandle, strValue, 0&, 0&, ByVal strBuffer, lngLen) = ERROR_SUCCESS Then
                    If lngLen > 0 Then
                        ReDim bytArray(lngLen - 1)
                        bytArray = Left$(strBuffer, lngLen)
                        strBuffer = StrConv(bytArray, vbFromUnicode)
                        Erase bytArray
                        If Left$(strBuffer, 4) = "\??\" Then
                            strBuffer = Mid$(strBuffer, 5, InStr(1, strBuffer, "{") - 6)
                            GetDeviceInstance = Replace(strBuffer, "#", "\")
                        End If
                    End If
                End If
            End If
        End If
        RegCloseKey lngHandle
    End If
End Function

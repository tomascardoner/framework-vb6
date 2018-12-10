Attribute VB_Name = "CSM_Registry"
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'API FUNCTIONS
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx_String Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx_Long Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
'Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As String, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteKeyEx Lib "advapi32.dll" Alias "RegDeleteKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal samDesired As Long, ByVal Reserved As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'ROOT KEYS
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

'KEY PERMISSIONS
Private Const READ_CONTROL = &H20000
Private Const READ_WRITE = 2
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_SET_VALUE = &H2
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_WOW64_64KEY = &H100
Private Const KEY_WOW64_32KEY = &H200

'KEY VOLATILITY
Private Const REG_OPTION_BACKUP_RESTORE = 4
Private Const REG_OPTION_CREATE_LINK = 2
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_RESERVED = 0
Private Const REG_OPTION_VOLATILE = 1

'ERROR CODES
Private Const ERROR_SUCCESS = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

'DATA TYPES
Private Const REG_BINARY = 3
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_RESOURCE_LIST = 8
Private Const REG_SZ = 1

'ENUMS
Public Enum CSRegistryKeys
    csrkHKEY_CLASSES_ROOT = HKEY_CLASSES_ROOT
    csrkHKEY_CURRENT_USER = HKEY_CURRENT_USER
    csrkHKEY_LOCAL_MACHINE = HKEY_LOCAL_MACHINE
    csrkHKEY_USERS = HKEY_USERS
    csrkHKEY_PERFORMANCE_DATA = HKEY_PERFORMANCE_DATA
    csrkHKEY_CURRENT_CONFIG = HKEY_CURRENT_CONFIG
    csrkHKEY_DYN_DATA = HKEY_DYN_DATA
End Enum

Public Enum CSRegistryDataTypes
    csrdtString
    csrdtNumberInteger
    csrdtNumberDecimal
    csrdtCurrency
    csrdtDateTime
    csrdtBoolean
End Enum

'*****************************************************************************************
'GET VALUE
'*****************************************************************************************
Public Function GetValue(ByVal Key As CSRegistryKeys, ByVal SubKey As String, ByVal ValueName As String, ByVal DefaultValue As Variant, ByVal DataType As CSRegistryDataTypes) As Variant
    Dim ValueString As String
    Dim ValueLong As Long
    Dim ReturnedValue As Long
    Dim hKey As Long
    Dim ValueType As Long
    Dim ValueSize As Long
    
    ReturnedValue = RegOpenKeyEx(Key, SubKey, 0, KEY_QUERY_VALUE, hKey)
    If ReturnedValue = ERROR_SUCCESS Then
        ValueString = String(255, vbNullChar)
        ValueSize = Len(ValueString)
        ReturnedValue = RegQueryValueEx_String(hKey, ValueName, CLng(0), ValueType, ValueString, ValueSize)
        If ReturnedValue = ERROR_SUCCESS Then
            Select Case ValueType
                Case REG_DWORD
                    ReturnedValue = RegQueryValueEx_Long(hKey, ValueName, CLng(0), ValueType, ValueLong, ValueSize)
                    RegCloseKey hKey
                    GetValue = ValueLong
                    Exit Function
                Case REG_SZ, REG_BINARY
                    If ValueSize > 0 Then
                        ValueString = Left(ValueString, ValueSize - 1)
                    End If
            End Select
        End If
        RegCloseKey hKey
    End If

    On Error Resume Next
    
    Select Case DataType
        Case csrdtString
            If ReturnedValue = ERROR_SUCCESS And ValueSize > 0 Then
                GetValue = CStr(ValueString)
            Else
                GetValue = CStr(DefaultValue)
            End If
            
        Case csrdtNumberInteger
            If ReturnedValue = ERROR_SUCCESS And ValueSize > 0 And IsNumeric(ValueString) Then
                GetValue = CLng(ValueString)
            Else
                GetValue = CLng(DefaultValue)
            End If
            
        Case csrdtNumberDecimal
            If ReturnedValue = ERROR_SUCCESS And ValueSize > 0 And IsNumeric(ValueString) Then
                GetValue = CDbl(ValueString)
            Else
                GetValue = CDbl(DefaultValue)
            End If
            
        Case csrdtCurrency
            If ReturnedValue = ERROR_SUCCESS And ValueSize > 0 And IsNumeric(ValueString) Then
                GetValue = CCur(ValueString)
            Else
                GetValue = CCur(DefaultValue)
            End If
            
        Case csrdtDateTime
            If ReturnedValue = ERROR_SUCCESS And ValueSize > 0 And IsDate(ValueString) Then
                GetValue = CDate(ValueString)
            Else
                GetValue = CDate(DefaultValue)
            End If
            
        Case csrdtBoolean
            If ReturnedValue = ERROR_SUCCESS And ValueSize > 0 Then
                If IsNumeric(ValueString) Then
                    GetValue = CBool(ValueString <> 0)
                ElseIf ValueString = "False" Or ValueString = "Falso" Or ValueString = "no" Then
                    GetValue = False
                ElseIf ValueString = "True" Or ValueString = "Verdadero" Or ValueString = "yes" Then
                    GetValue = True
                Else
                    GetValue = CBool(DefaultValue)
                End If
            Else
                GetValue = CBool(DefaultValue)
            End If
    End Select
End Function

Public Function GetValue_FromApplication_LocalMachine(ByVal SubKey As String, ByVal ValueName As String, ByVal DefaultValue As Variant, ByVal DataType As CSRegistryDataTypes) As Variant
    GetValue_FromApplication_LocalMachine = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey, ValueName, DefaultValue, DataType)
End Function

Public Function GetValue_FromApplication_CurrentUser(ByVal SubKey As String, ByVal ValueName As String, ByVal DefaultValue As Variant, ByVal DataType As CSRegistryDataTypes) As Variant
    GetValue_FromApplication_CurrentUser = CSM_Registry.GetValue(csrkHKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey, ValueName, DefaultValue, DataType)
End Function


'*****************************************************************************************
'DELETE VALUE
'*****************************************************************************************

Public Function DeleteValue(ByVal Key As CSRegistryKeys, ByVal SubKey As String, ByVal ValueName As String) As Boolean
    Dim ReturnedValue As Long
    Dim hKey As Long
    
    ReturnedValue = RegOpenKeyEx(Key, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If ReturnedValue = ERROR_SUCCESS Then
        ReturnedValue = RegDeleteValue(hKey, ValueName)
        RegCloseKey hKey
    End If
End Function

Public Function DeleteValue_FromApplication_LocalMachine(ByVal SubKey As String, ByVal ValueName As String) As Boolean
    DeleteValue_FromApplication_LocalMachine = CSM_Registry.DeleteValue(csrkHKEY_LOCAL_MACHINE, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey, ValueName)
End Function

Public Function DeleteValue_FromApplication_CurrentUser(ByVal SubKey As String, ByVal ValueName As String) As Boolean
    DeleteValue_FromApplication_CurrentUser = CSM_Registry.DeleteValue(csrkHKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey, ValueName)
End Function


'*****************************************************************************************
'SET VALUE
'*****************************************************************************************

Public Function SetValue(ByVal Key As CSRegistryKeys, ByVal SubKey As String, ByVal ValueName As String, ByVal Value As Variant) As Boolean
    Dim ReturnedValue As Long
    Dim hKey As Long
    Dim StringedValue As String
    
    'ADD THE NULL CHAR TO TEH END OF THE STRING
    StringedValue = CStr(Value) & vbNullChar
    
RESTART:
    ReturnedValue = RegOpenKeyEx(Key, SubKey, 0, KEY_ALL_ACCESS, hKey)
    Select Case ReturnedValue
        Case ERROR_SUCCESS
            'KEY SUCCESFULY OPEN
            ReturnedValue = RegSetValueEx(hKey, ValueName, CLng(0), REG_SZ, StringedValue, Len(StringedValue))
            Select Case ReturnedValue
                Case ERROR_SUCCESS
                    SetValue = True
                Case Else
                    Debug.Print CSM_Registry.GetErrorMessage(ReturnedValue)
            End Select
            RegCloseKey hKey
        Case ERROR_BADKEY
            'KEY DOESN'T EXIST, SO CREATE IT
            RegCreateKeyEx Key, SubKey, CLng(0), vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, CLng(0), CLng(0), CLng(0)
            GoTo RESTART
        Case Else
            Debug.Print CSM_Registry.GetErrorMessage(ReturnedValue)
    End Select
End Function

Public Function SetValue_ToApplication_LocalMachine(ByVal SubKey As String, ByVal ValueName As String, ByVal Value As Variant) As Boolean
    SetValue_ToApplication_LocalMachine = CSM_Registry.SetValue(csrkHKEY_LOCAL_MACHINE, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey, ValueName, Value)
End Function

Public Function SetValue_ToApplication_CurrentUser(ByVal SubKey As String, ByVal ValueName As String, ByVal Value As Variant) As Boolean
    SetValue_ToApplication_CurrentUser = CSM_Registry.SetValue(csrkHKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey, ValueName, Value)
End Function

'*****************************************************************************************
'ENUMERATE SUBKEYS
'*****************************************************************************************
Public Function EnumerateSubKeys(ByVal Key As CSRegistryKeys, ByVal SubKey As String) As Collection
    Dim ReturnedValue As Long
    Dim hKey As Long
    Dim ValueIndex As Long
    Dim ValueName As String
    Dim ValueNameLenght As Long
    
    Set EnumerateSubKeys = New Collection

    ReturnedValue = RegOpenKeyEx(Key, SubKey, 0, KEY_ENUMERATE_SUB_KEYS Or KEY_WOW64_64KEY, hKey)
    If ReturnedValue = ERROR_SUCCESS Then
        Do While ReturnedValue = ERROR_SUCCESS
            ValueName = String(255, vbNullChar)
            ValueNameLenght = Len(ValueName)
            ReturnedValue = RegEnumKey(hKey, ValueIndex, ValueName, ValueNameLenght)
            If ReturnedValue = ERROR_SUCCESS Then
                ValueNameLenght = InStr(1, ValueName, vbNullChar) - 1
                ValueName = Left(ValueName, ValueNameLenght)
                
                ValueIndex = ValueIndex + 1
                
                EnumerateSubKeys.Add ValueName
            End If
        Loop
        
        RegCloseKey hKey
    End If
End Function

Public Function EnumerateSubKeys_FromApplication_LocalMachine(ByVal SubKey As String) As Collection
    Set EnumerateSubKeys_FromApplication_LocalMachine = CSM_Registry.EnumerateSubKeys(csrkHKEY_LOCAL_MACHINE, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey)
End Function

Public Function EnumerateSubKeys_FromApplication_CurrentUser(ByVal SubKey As String) As Collection
    Set EnumerateSubKeys_FromApplication_CurrentUser = CSM_Registry.EnumerateSubKeys(csrkHKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey)
End Function

'*****************************************************************************************
'ENUMERATE VALUES
'*****************************************************************************************
Public Function EnumerateValues(ByVal Key As CSRegistryKeys, ByVal SubKey As String, ByRef CValuesNames As Collection, ByRef CValuesTypes As Collection, ByRef CValues As Collection, ByVal FilterPrefix As String, ByVal RemovePrefix As Boolean, ParamArray avExcept() As Variant) As Boolean
    Dim ReturnedValue As Long
    Dim hKey As Long
    Dim ValueIndex As Long
    Dim ValueName As String
    Dim ValueNameLenght As Long
    Dim ValueType As Long
    Dim Value As String
    Dim ValueLenght As Long
    Dim ValueCharIndex As Long
    Dim ValueTemp As String
    Dim ExceptValueName As Variant
    Dim ExceptFound As Boolean

    Set CValuesNames = New Collection
    Set CValuesTypes = New Collection
    Set CValues = New Collection
    
    ReturnedValue = RegOpenKeyEx(Key, SubKey, 0, KEY_QUERY_VALUE, hKey)
    If ReturnedValue = ERROR_SUCCESS Then
        Do While ReturnedValue = ERROR_SUCCESS
            ValueName = String(255, vbNullChar)
            ValueNameLenght = Len(ValueName)
            Value = String(32767, vbNullChar)
            ValueLenght = Len(Value)
            ReturnedValue = RegEnumValue(hKey, ValueIndex, ValueName, ValueNameLenght, 0, ValueType, ByVal Value, ValueLenght)
            If ReturnedValue = ERROR_SUCCESS Then
                ValueIndex = ValueIndex + 1
                ValueName = Left(ValueName, ValueNameLenght)
                
                ExceptFound = False
                If UBound(avExcept) > 0 Then
                    For Each ExceptValueName In avExcept
                        If ValueName = ExceptValueName Then
                            ExceptFound = True
                            Exit For
                        End If
                    Next ExceptValueName
                End If
                
                If Not ExceptFound Then
                    If Left(ValueName, Len(FilterPrefix)) = FilterPrefix Then
                        If RemovePrefix Then
                            ValueName = Mid(ValueName, Len(FilterPrefix) + 1)
                        End If
                        CValuesNames.Add ValueName, CSM_Constant.KEY_STRINGER & ValueName
                        CValuesTypes.Add ValueType, CSM_Constant.KEY_STRINGER & ValueName
                        Select Case ValueType
                            Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                                Value = Left(Value, ValueLenght - 1)
                            Case REG_BINARY
                                ValueTemp = ""
                                For ValueCharIndex = 1 To ValueLenght
                                    ValueTemp = ValueTemp & IIf(Asc(Mid(Value, ValueCharIndex, 1)) <= 15, "0", "") & Hex(Asc(Mid(Value, ValueCharIndex, 1))) & " "
                                Next ValueCharIndex
                                If ValueLenght = 0 Then
                                    Value = ""
                                Else
                                    Value = Left(ValueTemp, Len(ValueTemp) - 1)
                                End If
                            Case REG_DWORD
                                ValueTemp = ""
                                For ValueCharIndex = ValueLenght To 1 Step -1
                                    ValueTemp = ValueTemp & IIf(Asc(Mid(Value, ValueCharIndex, 1)) <= 15, "0", "") & Hex(Asc(Mid(Value, ValueCharIndex, 1)))
                                Next ValueCharIndex
                                Value = CSM_Conversion.FromBase(ValueTemp, 16)
                            Case Else
                                Stop
                        End Select
                        CValues.Add Value, CSM_Constant.KEY_STRINGER & ValueName
                    End If
                End If
            End If
        Loop
        
        RegCloseKey hKey
        
        EnumerateValues = True
    End If
End Function

Public Function EnumerateValues_FromApplication_LocalMachine(ByVal SubKey As String, ByRef CValuesNames As Collection, ByRef CValuesTypes As Collection, ByRef CValues As Collection, ByVal FilterPrefix As String, ByVal RemovePrefix As Boolean, ParamArray avExcept() As Variant) As Boolean
    EnumerateValues_FromApplication_LocalMachine = CSM_Registry.EnumerateValues(csrkHKEY_LOCAL_MACHINE, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey, CValuesNames, CValuesTypes, CValues, FilterPrefix, RemovePrefix, avExcept)
End Function

Public Function EnumerateValues_FromApplication_CurrentUser(ByVal SubKey As String, ByRef CValuesNames As Collection, ByRef CValuesTypes As Collection, ByRef CValues As Collection, ByVal FilterPrefix As String, ByVal RemovePrefix As Boolean, ParamArray avExcept() As Variant) As Boolean
    EnumerateValues_FromApplication_CurrentUser = CSM_Registry.EnumerateValues(csrkHKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey, CValuesNames, CValuesTypes, CValues, FilterPrefix, RemovePrefix, avExcept)
End Function

'*****************************************************************************************
'DELETE SUBKEY
'*****************************************************************************************
Public Function DeleteSubKey(ByVal Key As CSRegistryKeys, ByVal SubKey As String, Optional DeleteChildSubKeys As Boolean = False) As Boolean
    Dim ReturnedValue As Long
    Dim CChildSubKeys As Collection
    Dim ChildSubKey As Variant
    Static OperatingSystemBits As Byte
    
    If OperatingSystemBits = 0 Then
        If CSM_System.IsHost64Bit Then
            OperatingSystemBits = 64
        Else
            OperatingSystemBits = 32
        End If
    End If
    
    If DeleteChildSubKeys Then
        Set CChildSubKeys = EnumerateSubKeys(Key, SubKey)
        For Each ChildSubKey In CChildSubKeys
            If Not DeleteSubKey(Key, SubKey & "\" & ChildSubKey, True) Then
                Exit For
            End If
        Next ChildSubKey
    End If
    
    If OperatingSystemBits = 64 Then
        ReturnedValue = RegDeleteKeyEx(Key, SubKey, KEY_ALL_ACCESS Or KEY_WOW64_64KEY, 0)
    Else
        ReturnedValue = RegDeleteKey(Key, SubKey)
    End If
    
    DeleteSubKey = (ReturnedValue = 0)
End Function

Public Function DeleteSubKeyFromPath(ByVal Path As String) As Boolean
    Dim KeyIndex As CSRegistryKeys
    
    KeyIndex = GetKeyIndex(Path)
    If KeyIndex <> 0 Then
        DeleteSubKeyFromPath = DeleteSubKey(KeyIndex, CSM_String.RemoveFirstSubString(Path, "\"), True)
    End If
End Function

Public Function DeleteSubKey_FromApplication_LocalMachine(ByVal SubKey As String) As Boolean
    DeleteSubKey_FromApplication_LocalMachine = CSM_Registry.DeleteSubKey(csrkHKEY_LOCAL_MACHINE, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey)
End Function

Public Function DeleteSubKey_FromApplication_CurrentUser(ByVal SubKey As String) As Boolean
    DeleteSubKey_FromApplication_CurrentUser = CSM_Registry.DeleteSubKey(csrkHKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & App.ProductName & IIf(SubKey <> "", "\", "") & SubKey)
End Function

'*****************************************************************************************
'GET ERROR MESSAGE
'*****************************************************************************************

Private Function GetErrorMessage(ByVal ErrorNumber As Long)
    Select Case ErrorNumber
        Case ERROR_SUCCESS = 0
            GetErrorMessage = ""
        Case ERROR_BADDB = 1
            GetErrorMessage = "El registro se encuentra dañado."
        Case ERROR_BADKEY = 2
            GetErrorMessage = "No se encontró la clave del registro especificada."
        Case ERROR_CANTOPEN = 3
            GetErrorMessage = "No se pudo abrir la clave del registro especificada."
        Case ERROR_CANTREAD = 4
            GetErrorMessage = "No se pudo leer el valor del registro especificado."
        Case ERROR_CANTWRITE = 5
            GetErrorMessage = "No se pudo escribir el valor del registro especificado."
        Case ERROR_OUTOFMEMORY = 6
            GetErrorMessage = "Memoria insuficiente accediendo al registro."
        Case ERROR_ARENA_TRASHED = 7
            GetErrorMessage = "Se econtró 'basura' accediendo al registro."
        Case ERROR_ACCESS_DENIED = 8
            GetErrorMessage = "Se ha denegado el acceso al registro."
        Case ERROR_INVALID_PARAMETERS = 87
            GetErrorMessage = "Los parámetros especificados para acceder al registro, no son válidos."
        Case ERROR_NO_MORE_ITEMS = 259
            GetErrorMessage = "No hay más items para leer en el registro."
    End Select
End Function

Public Function GetKeyIndex(ByVal Path As String) As CSRegistryKeys
    Select Case CSM_String.GetSubString(Path, 1, "\")
        Case "HKEY_CLASSES_ROOT"
            GetKeyIndex = csrkHKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER"
            GetKeyIndex = csrkHKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            GetKeyIndex = csrkHKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            GetKeyIndex = csrkHKEY_USERS
        Case "HKEY_CURRENT_CONFIG"
            GetKeyIndex = csrkHKEY_CURRENT_CONFIG
        Case Else
            GetKeyIndex = 0
    End Select
End Function

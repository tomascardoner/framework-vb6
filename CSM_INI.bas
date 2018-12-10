Attribute VB_Name = "CSM_INI"
Option Explicit

Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Enum CSINIDataTypes
    csidtString
    csidtNumberInteger
    csidtNumberDecimal
    csidtCurrency
    csidtDateTime
    csidtBoolean
End Enum

Public Function GetSectionsNames(ByVal FileName As String, Optional ByVal FilterPrefix As String = "") As Collection
    Dim Value As String
    Dim ValueSize As Long
    Dim ReturnedValue As Long
    Dim aValues() As String
    Dim Index As Long
    
    Set GetSectionsNames = New Collection
    Value = String(100000000, vbNullChar)
    ValueSize = Len(Value)
    ReturnedValue = GetPrivateProfileSectionNames(Value, ValueSize, FileName)
    If ReturnedValue > 0 Then
        Value = Left(Value, ReturnedValue)
        If Right(Value, 1) = vbNullChar Then
            Value = Left(Value, ReturnedValue - 1)
        End If
        aValues = Split(Value, vbNullChar)
        On Error Resume Next
        For Index = 0 To UBound(aValues)
            If Left(aValues(Index), Len(FilterPrefix)) = FilterPrefix Then
                GetSectionsNames.Add aValues(Index), KEY_STRINGER & aValues(Index)
            End If
        Next Index
    End If
End Function

Public Function GetSectionsNames_FromApplication(Optional ByVal FilterPrefix As String = "") As Collection
    Set GetSectionsNames_FromApplication = GetSectionsNames(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & ".ini", FilterPrefix)
End Function

Public Function GetSectionValues(ByVal FileName As String, ByVal Section As String) As Collection
    Dim Value As String
    Dim ValueSize As Long
    Dim ReturnedValue As Long
    Dim aValues() As String
    Dim Index As Long
    
    Set GetSectionValues = New Collection
    Value = String(100000000, vbNullChar)
    ValueSize = Len(Value)
    ReturnedValue = GetPrivateProfileSection(Section, Value, ValueSize, FileName)
    If ReturnedValue > 0 Then
        Value = Left(Value, ReturnedValue)
        If Right(Value, 1) = vbNullChar Then
            Value = Left(Value, ReturnedValue - 1)
        End If
        aValues = Split(Value, vbNullChar)
        For Index = 0 To UBound(aValues)
            GetSectionValues.Add aValues(Index)
        Next Index
    End If
End Function

Public Function GetValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Variant, ByVal DataType As CSINIDataTypes) As Variant
    Dim Value As String
    Dim ValueSize As Long
    Dim ReturnedValue As Long
    
    Value = String(255, vbNullChar)
    ValueSize = Len(Value)
    ReturnedValue = GetPrivateProfileString(Section, Key, "", Value, ValueSize, FileName)
    If ReturnedValue > 0 Then
        Value = Left(Value, ReturnedValue)
    End If

    On Error Resume Next
    
    Select Case DataType
        Case csidtString
            If ReturnedValue > 0 Then
                GetValue = CStr(Value)
            Else
                GetValue = CStr(DefaultValue)
            End If
            
        Case csidtNumberInteger
            If ReturnedValue > 0 And IsNumeric(Value) Then
                GetValue = CLng(Value)
            Else
                GetValue = CLng(DefaultValue)
            End If
            
        Case csidtNumberDecimal
            If ReturnedValue > 0 And IsNumeric(Value) Then
                GetValue = CDbl(Value)
            Else
                GetValue = CDbl(DefaultValue)
            End If
            
        Case csidtCurrency
            If ReturnedValue > 0 And IsNumeric(Value) Then
                GetValue = CCur(Value)
            Else
                GetValue = CCur(DefaultValue)
            End If
            
        Case csidtDateTime
            If ReturnedValue > 0 And IsDate(Value) Then
                GetValue = CDate(Value)
            Else
                GetValue = CDate(DefaultValue)
            End If
            
        Case csidtBoolean
            If ReturnedValue > 0 Then
                If IsNumeric(Value) Then
                    GetValue = CBool(Value <> 0)
                ElseIf Value = "False" Or Value = "Falso" Then
                    GetValue = False
                ElseIf Value = "True" Or Value = "Verdadero" Then
                    GetValue = True
                Else
                    GetValue = CBool(DefaultValue)
                End If
            Else
                GetValue = CBool(DefaultValue)
            End If
    End Select
End Function

Public Function GetValue_FromApplication(ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Variant, ByVal DataType As CSINIDataTypes) As Variant
    GetValue_FromApplication = GetValue(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & ".ini", Section, Key, DefaultValue, DataType)
End Function

Public Function SetValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
    SetValue = (WritePrivateProfileString(Section, Key, Value, FileName) = 0)
End Function

Public Function SetValue_ToApplication(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
    SetValue_ToApplication = SetValue(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & ".ini", Section, Key, Value)
End Function

Attribute VB_Name = "CSM_Session"
Option Explicit

'API FUNCTIONS
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SetComputerNameAPI Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function GetUserNameAPI Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Public Function GetComputerName() As String
    Dim BufferLenght As Long
    Dim Buffer As String
    Dim ReturnedValue As Long
    Static ComputerName As String
    
    Const MAX_COMPUTERNAME_LENGTH As Long = 31
    
    If ComputerName = "" Then
        BufferLenght = MAX_COMPUTERNAME_LENGTH + 1
        Buffer = String(BufferLenght, vbNullChar)
        ReturnedValue = GetComputerNameAPI(Buffer, BufferLenght)
        If ReturnedValue <> 0 Then
            ComputerName = Left(Buffer, BufferLenght)
        End If
    End If
    GetComputerName = ComputerName
End Function

Public Function SetComputerName(ByVal ComputerName As String) As Boolean
    SetComputerName = CBool(SetComputerNameAPI(ComputerName))
End Function

Public Function GetUserName() As String
    Dim BufferLenght As Long
    Dim Buffer As String
    Dim ReturnedValue As Long
    
    BufferLenght = 255
    Buffer = String(BufferLenght, vbNullChar)
    ReturnedValue = GetUserNameAPI(Buffer, BufferLenght)
    If ReturnedValue <> 0 Then
        GetUserName = Left(Buffer, BufferLenght - 1)
    End If
End Function

Public Function GetNetworkUserName() As String
    Dim BufferLenght As Long
    Dim Buffer As String
    Dim ReturnedValue As Long
    
    BufferLenght = 255
    Buffer = String(BufferLenght, vbNullChar)
    ReturnedValue = WNetGetUser(vbNullString, Buffer, BufferLenght)
    If ReturnedValue = 0 Then
        GetNetworkUserName = Left(Buffer, InStr(Buffer, vbNullChar) - 1)
    End If
End Function

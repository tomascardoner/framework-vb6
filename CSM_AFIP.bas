Attribute VB_Name = "CSM_AFIP"
Option Explicit

Public Function DigitoVerificadorCUIT(ByVal CUIT As String) As Byte
    Dim XA As Integer
    Dim XB As Integer
    Dim XC As Integer
    Dim XD As Integer
    Dim XE As Integer
    Dim XF As Integer
    Dim XG As Integer
    Dim XH As Integer
    Dim XI As Integer
    Dim XJ As Integer
    Dim x As Integer

    'Verifica si el tamaño es el correcto.
    If Len(CUIT) = 10 Then
        'Individualiza y multiplica los dígitos.
        XA = Val(Mid$(CUIT, 1, 1)) * 5
        XB = Val(Mid$(CUIT, 2, 1)) * 4
        XC = Val(Mid$(CUIT, 3, 1)) * 3
        XD = Val(Mid$(CUIT, 4, 1)) * 2
        XE = Val(Mid$(CUIT, 5, 1)) * 7
        XF = Val(Mid$(CUIT, 6, 1)) * 6
        XG = Val(Mid$(CUIT, 7, 1)) * 5
        XH = Val(Mid$(CUIT, 8, 1)) * 4
        XI = Val(Mid$(CUIT, 9, 1)) * 3
        XJ = Val(Mid$(CUIT, 10, 1)) * 2
    
        'Suma los resultantes.
        x = XA + XB + XC + XD + XE + XF + XG + XH + XI + XJ
    
        'Calcula el dígito de control.
        DigitoVerificadorCUIT = (11 - (x Mod 11)) Mod 11
    End If
End Function

Public Function VerificarCUIT_SinGuiones(ByVal CUIT As String) As Boolean
    If Len(CUIT) = 11 Then
        VerificarCUIT_SinGuiones = (DigitoVerificadorCUIT(Left(CUIT, 10)) = Val(Mid$(CUIT, 11, 1)))
    End If
End Function

Public Function VerificarCUIT_ConGuiones(ByVal CUIT As String) As Boolean
    If Len(CUIT) = 13 Then
        If Mid(CUIT, 3, 1) = "-" And Mid(CUIT, 12, 1) = "-" Then
            CUIT = Mid(CUIT, 1, 2) & Mid(CUIT, 4, 8) & Mid(CUIT, 13, 1)
            VerificarCUIT_ConGuiones = (DigitoVerificadorCUIT(Left(CUIT, 10)) = Val(Mid$(CUIT, 11, 1)))
        End If
    End If
End Function

Public Function VerificarCUIT(ByVal CUIT As String) As Boolean
    Dim Prefix As String
    Dim Count As Long
    
    'PRIMERO PREPARO EL STRING
    If Len(CUIT) = 13 Then
        'TIENE 13 CARACTERES, VERIFICO QUE TENGA LOS GUIONES EN EL LUGAR CORRECTO
        If Mid(CUIT, 3, 1) = "-" And Mid(CUIT, 12, 1) = "-" Then
            'LIMPIO EL STRING DE LOS CARACTERES NO NUMERICOS
            CUIT = CSM_String.CleanNotNumericChars(CUIT)
        End If
    ElseIf Len(CUIT) = 11 Then
        'TIENE 11 CARACTERES, LIMPIO EL STRING DE LOS CARACTERES NO NUMERICOS
        CUIT = CSM_String.CleanNotNumericChars(CUIT)
    End If
    
    'VERIFICO QUE TENGA 11 CARACTERES. YA SE QUE SON NUMEROS POR LOS PASOS ANTERIORES
    If Len(CUIT) = 11 Then
        Prefix = Left(CUIT, 2)
        If Prefix = "20" Or Prefix = "23" Or Prefix = "24" Or Prefix = "27" Or Prefix = "30" Or Prefix = "33" Or Prefix = "34" Then
            Count = Mid(CUIT, 1, 1) * 5 + Mid(CUIT, 2, 1) * 4 + Mid(CUIT, 3, 1) * 3 + Mid(CUIT, 4, 1) * 2 + Mid(CUIT, 5, 1) * 7 + Mid(CUIT, 6, 1) * 6 + Mid(CUIT, 7, 1) * 5 + Mid(CUIT, 8, 1) * 4 + Mid(CUIT, 9, 1) * 3 + Mid(CUIT, 10, 1) * 2 + Mid(CUIT, 11, 1) * 1
            VerificarCUIT = ((Count Mod 11) = 0)
        End If
    End If
End Function

Public Function DatabaseFileFixAndUnfix(ByVal FixFile As Boolean, ByVal FileName As String, ByVal MakeBackup As Boolean) As Boolean
    Dim FileNumber As Integer
    Dim Position As Long
    Dim Value As Byte
    
    Const OFFSET As Long = 10
    Const XOR_VALUE As Long = 69
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If MakeBackup Then
        FileSystem.FileCopy FileName, FileName & "__backup"
    End If
    
    FileNumber = FreeFile()
    Open FileName For Random Access Read Write Lock Read Write As #FileNumber Len = 1
    
    Get #FileNumber, 1, Value
    If (FixFile And Value = XOR_VALUE) Or (FixFile = False And Value = 0) Then
        For Position = 1 To 391 Step OFFSET
            Get #FileNumber, Position, Value
            Value = Value Xor XOR_VALUE
            Put #FileNumber, Position, Value
        Next Position
    End If
    
    Close #FileNumber
    
    DatabaseFileFixAndUnfix = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Modules.CSM_AFIP.DatabaseFileFixAndUnfix", "Error al reparar la Base de Datos AFIP."
    On Error Resume Next
    Close #FileNumber
End Function


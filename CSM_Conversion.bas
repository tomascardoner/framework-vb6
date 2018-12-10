Attribute VB_Name = "CSM_Conversion"
Option Explicit

'Hexadecimal Constants
Private Const BIN_0 As String = "0000"
Private Const BIN_1 As String = "0001"
Private Const BIN_2 As String = "0010"
Private Const BIN_3 As String = "0011"
Private Const BIN_4 As String = "0100"
Private Const BIN_5 As String = "0101"
Private Const BIN_6 As String = "0110"
Private Const BIN_7 As String = "0111"
Private Const BIN_8 As String = "1000"
Private Const BIN_9 As String = "1001"
Private Const BIN_A As String = "1010"
Private Const BIN_B As String = "1011"
Private Const BIN_C As String = "1100"
Private Const BIN_D As String = "1101"
Private Const BIN_E As String = "1110"
Private Const BIN_F As String = "1111"

Public Function CharHex2Bin(ByVal Value As String) As String
    Select Case UCase(Value)
        Case "0"
            CharHex2Bin = BIN_0
        Case "1"
            CharHex2Bin = BIN_1
        Case "2"
            CharHex2Bin = BIN_2
        Case "3"
            CharHex2Bin = BIN_3
        Case "4"
            CharHex2Bin = BIN_4
        Case "5"
            CharHex2Bin = BIN_5
        Case "6"
            CharHex2Bin = BIN_6
        Case "7"
            CharHex2Bin = BIN_7
        Case "8"
            CharHex2Bin = BIN_8
        Case "9"
            CharHex2Bin = BIN_9
        Case "A"
            CharHex2Bin = BIN_A
        Case "B"
            CharHex2Bin = BIN_B
        Case "C"
            CharHex2Bin = BIN_C
        Case "D"
            CharHex2Bin = BIN_D
        Case "E"
            CharHex2Bin = BIN_E
        Case "F"
            CharHex2Bin = BIN_F
    End Select
End Function

Public Function CharBin2Hex(ByVal BinaryText As String) As String
    Select Case BinaryText
        Case BIN_0
            CharBin2Hex = "0"
        Case BIN_1
            CharBin2Hex = "1"
        Case BIN_2
            CharBin2Hex = "2"
        Case BIN_3
            CharBin2Hex = "3"
        Case BIN_4
            CharBin2Hex = "4"
        Case BIN_5
            CharBin2Hex = "5"
        Case BIN_6
            CharBin2Hex = "6"
        Case BIN_7
            CharBin2Hex = "7"
        Case BIN_8
            CharBin2Hex = "8"
        Case BIN_9
            CharBin2Hex = "9"
        Case BIN_A
            CharBin2Hex = "A"
        Case BIN_B
            CharBin2Hex = "B"
        Case BIN_C
            CharBin2Hex = "C"
        Case BIN_D
            CharBin2Hex = "D"
        Case BIN_E
            CharBin2Hex = "E"
        Case BIN_F
            CharBin2Hex = "F"
        Case Else
    End Select
End Function

Public Function Hex2Bin(ByVal BinaryText As String) As String
    Dim CharPos As Long
    
    For CharPos = 1 To Len(BinaryText)
        Hex2Bin = Hex2Bin & CharHex2Bin(Mid(BinaryText, CharPos, 1))
    Next CharPos
End Function

Public Function Bin2Hex(ByVal BinaryText As String) As String
    Dim CharPos As Long
    
    For CharPos = 1 To Len(BinaryText) Step 4
        Bin2Hex = Bin2Hex & CharBin2Hex(Mid(BinaryText, CharPos, 4))
    Next CharPos
End Function

Public Function Bin2Text(ByVal Value As String) As String
    Dim CharPos As Long
    
    For CharPos = 1 To Len(Value) Step 8
        Bin2Text = Bin2Text & Chr(FromBase(Mid(Value, CharPos, 8), 2))
    Next CharPos
End Function

Public Function FromBase(ByVal BaseNumber As String, ByVal OldBase As Integer) As Double
    Dim CharPos As Long
    Dim LetterVal As Integer
    
    On Error Resume Next

    For CharPos = 1 To Len(BaseNumber)
        LetterVal = Asc(Mid(BaseNumber, Len(BaseNumber) - CharPos + 1, 1)) - 48
        If LetterVal > 9 Then
            LetterVal = LetterVal - 7
        End If
        If LetterVal > OldBase Then
            GoTo InvalidNumber
        End If
        FromBase = FromBase + ((OldBase ^ (CharPos - 1)) * LetterVal)
    Next CharPos
    
    Exit Function
    
InvalidNumber:
    FromBase = 0
End Function

Public Function ToBase(ByVal DecimalNumber As Integer, NewBase As Integer) As String
    Dim ModBase As Double
    
    Do
        ModBase = CDbl(DecimalNumber - (Int(DecimalNumber / NewBase)) * NewBase)
        DecimalNumber = Int(DecimalNumber / NewBase)
        If ModBase > 9 Then
            ModBase = ModBase + 7
        End If
        ToBase = Chr(ModBase + 48) & ToBase
        
    Loop Until DecimalNumber = 0
End Function

Public Function DisplayBinaryData(ByVal Data As String) As String
    Dim HexNumberPosition As Long
    Dim HexNumber As String
    Dim DecimalNumber As Long
    Dim Counter As Long
    Dim StringedData As String
    
    If Len(Data) > 0 Then
        For HexNumberPosition = 1 To Len(Data) Step 3
            If Counter = 0 Then
                DisplayBinaryData = DisplayBinaryData & Format(((HexNumberPosition - 1) / 3), "0000") & " | "
                StringedData = ""
            End If
            Counter = Counter + 1
            HexNumber = Mid(Data, HexNumberPosition, 2)
            DisplayBinaryData = DisplayBinaryData & HexNumber & " "
            DecimalNumber = FromBase(HexNumber, 16)
            If DecimalNumber < 32 Then
                StringedData = StringedData & "."
            Else
                StringedData = StringedData & Chr(DecimalNumber)
            End If
            If Counter = 8 Then
                DisplayBinaryData = DisplayBinaryData & "| " & StringedData & vbCrLf
                Counter = 0
            End If
        Next HexNumberPosition
        If Counter > 0 Then
            DisplayBinaryData = DisplayBinaryData & String((8 - Counter) * 3, " ") & "| " & StringedData
        End If
    End If
End Function

Public Function BinaryData2Text(ByVal Data As String) As String
    Dim HexNumberPosition As Long
    Dim HexNumber As String
    Dim DecimalNumber As Long
    Dim Counter As Long
    
    If Len(Data) > 0 Then
        For HexNumberPosition = 1 To Len(Data) Step 3
            Counter = Counter + 1
            HexNumber = Mid(Data, HexNumberPosition, 2)
            DecimalNumber = FromBase(HexNumber, 16)
            BinaryData2Text = BinaryData2Text & Chr(DecimalNumber)
        Next HexNumberPosition
    End If
End Function

Public Function Text2Bin(ByVal Value As String) As String
    Dim CharPos As Long
    
    Value = Text2Hex(Value)
    For CharPos = 1 To Len(Value)
        Text2Bin = Text2Bin & CharHex2Bin(Mid(Value, CharPos, 1))
    Next CharPos
End Function

Public Function Hex2Dec(ByVal Value As String) As Long
    Hex2Dec = CLng("&H" & Value)
End Function

Public Function Text2Dec(ByVal Value As String) As Long
    Text2Dec = Hex2Dec(Text2Hex(Value))
End Function

Public Function Text2Hex(ByVal Value As String) As String
    Dim CharPos As Long
    Dim TempChar As String
    
    For CharPos = 1 To Len(Value)
        TempChar = Hex(Asc(Mid(Value, CharPos, 1)))
        Text2Hex = Text2Hex & IIf(Len(TempChar) = 1, "0" & TempChar, TempChar)
    Next CharPos
End Function


'////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////
'ARC & CIRCLES

Public Function Degrees2Radians(ByVal Value As Single) As Single
    Degrees2Radians = Value * (CSM_Math.PI / 180)
End Function

Public Function Radians2Degrees(ByVal Value As Single) As Single
    Radians2Degrees = Value * (180 / CSM_Math.PI)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_Encryption_DES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const INIT_VECTOR = "?n ãé‡"

Private Const TEXT_FILLER = "ôëÙÖ"
Private Const KEY_FILLER = "Š’KÝDÅ"

Friend Property Get PASSWORD_ENCRYPTION_KEY() As String
    PASSWORD_ENCRYPTION_KEY = "9A6B60C91AEAB43C682666E582B49EAB9A5F8F203C1161BD"
End Property

Friend Function EncryptString(ByVal Text As String, ByVal Key As String, ByVal UseTimer As Boolean) As String
    Dim EncryptedBlock As String
    Dim EncryptingBlock As String
    
    Dim TextLength As Long
    Dim KeyLength As Long
    Dim TimeString As String
    
    Dim BinaryText As String
    Dim BinaryKey As String
    
    Dim Counter As Long
    
    Dim TempInteger1 As Integer
    Dim TempInteger2 As Integer
    
    If Len(Trim(Text)) = 0 Or Len(Trim(Key)) = 0 Then
        Exit Function
    End If
    
    TextLength = Len(Text)
    KeyLength = Len(Key)
    
    '///////////////////////////////////////////////////////////////////////
    ' ADD HEADER DATA AS: USE OF TIMER, TEXT LENGTH, KEY LENGTH AND TIME
    '///////////////////////////////////////////////////////////////////////
    If UseTimer Then
        TimeString = Format(Date, "yyyymmdd") & Format(Timer * 100, "00000000")
        For Counter = 1 To 16 Step 2
            Text = Chr(Mid(TimeString, Counter, 2)) & Text
        Next Counter
    End If
    
    TempInteger1 = TextLength / 256
    TempInteger2 = TextLength Mod 256
    
    Text = IIf(UseTimer, "T", "F") & Chr(TempInteger1) & Chr(TempInteger2) & Text
    Text = Text & Left(TEXT_FILLER, IIf((TextLength Mod 8) = 0, 0, 8 - (TextLength Mod 8)))
    Key = Key & Left(KEY_FILLER, IIf((KeyLength Mod 8) = 0, 0, 8 - (KeyLength Mod 8)))
    
    'Prepare Key
    BinaryKey = CSM_Conversion.Text2Bin(Left(Key, 8))
    If Len(Key) > 8 Then
        For Counter = 9 To Len(Key) Step 8
            'Convert Key to Binary
            BinaryKey = ApplyXOR(BinaryKey, CSM_Conversion.Text2Bin(Mid(Key, Counter, 8)), 64)
        Next Counter
    End If
    '///////////////////////////////////////////////////////////////////////
    ' ENCRYPT INITIALIZATION VECTOR
    '///////////////////////////////////////////////////////////////////////
    'Convert Text
    BinaryText = CSM_Conversion.Text2Bin(INIT_VECTOR)
    
    EncryptedBlock = EncryptBinBlock(BinaryText, BinaryKey)
    '///////////////////////////////////////////////////////////////////////
    ' ENCRYPT EVERY BLOCK OF TEXT
    '///////////////////////////////////////////////////////////////////////
    For Counter = 1 To Len(Text) Step 8
        BinaryText = CSM_Conversion.Text2Bin(Mid(Text, Counter, 8))
    
        EncryptingBlock = ApplyXOR(EncryptedBlock, BinaryText, 64)
    
        EncryptedBlock = EncryptBinBlock(EncryptingBlock, BinaryKey)
        
        EncryptString = EncryptString & CSM_Conversion.Bin2Hex(EncryptedBlock)
    Next Counter
End Function

Private Function EncryptBinBlock(ByVal BinaryText As String, ByVal BinaryKey As String) As String
    Dim LoRoText As String
    Dim LoText As String
    Dim RoText As String
    
    Dim CoDoKey As String
    Dim CoKey As String
    Dim DoKey As String
    
    Dim ERiText As String
    Dim KiText As String
    Dim XORText As String
    Dim SBoxText As String
    Dim RiKiText As String
    Dim LiRiText As String
    
    Dim IPInverse As String
    
    Dim IterationNumber As Long
    Dim ShiftPosition As Long
        
    'Apply IP to Text
    LoRoText = ApplyIP(BinaryText)
    LoText = Left(LoRoText, 32)
    RoText = Mid(LoRoText, 33)

    'Apply PC1 to Key
    CoDoKey = ApplyPC1(BinaryKey)
    CoKey = Left(CoDoKey, 28)
    DoKey = Mid(CoDoKey, 29)
    
    'Do 16 Iterations to Apply E Expansion Function and PC2
    For IterationNumber = 1 To 16
        ERiText = ExpansionFunctionE(RoText)
        
        If IterationNumber = 1 Or IterationNumber = 2 Or IterationNumber = 9 Or IterationNumber = 16 Then
            ShiftPosition = 1
        Else
            ShiftPosition = 2
        End If
        
        CoKey = LeftShift(CoKey, ShiftPosition)
        DoKey = LeftShift(DoKey, ShiftPosition)
        CoDoKey = CoKey & DoKey
        
        KiText = ApplyPC2(CoDoKey)
        
        XORText = ApplyXOR(ERiText, KiText, 48)
        
        SBoxText = ApplySBoxes(XORText)
        
        RiKiText = ApplyP(SBoxText)
        
        LiRiText = ApplyXOR(LoText, RiKiText, 32)
        
        LoText = RoText
        RoText = LiRiText
    Next IterationNumber
    
    'IP Inverse for L16 and R16
    IPInverse = ApplyIPInverse(RoText & LoText)
    
    EncryptBinBlock = IPInverse
End Function

Friend Function DecryptString(ByVal Text As String, ByVal Key As String) As String
    Dim DecryptedBlock As String
    Dim DecryptingBlock As String
    Dim PreviousBlock As String
    
    Dim TextLength As Long
    Dim KeyLength As Long
    
    Dim BinaryText As String
    Dim BinaryKey As String
    
    Dim Counter As Long
    
    If Len(Trim(Text)) = 0 Or Len(Trim(Key)) = 0 Then
        Exit Function
    End If
    
    KeyLength = Len(Key)
    Key = Key & Left(KEY_FILLER, IIf((KeyLength Mod 8) = 0, 0, 8 - (KeyLength Mod 8)))
    
    'Prepare Key
    BinaryKey = CSM_Conversion.Text2Bin(Left(Key, 8))
    If Len(Key) > 8 Then
        For Counter = 9 To Len(Key) Step 8
            'Convert Key to Binary
            BinaryKey = ApplyXOR(BinaryKey, CSM_Conversion.Text2Bin(Mid(Key, Counter, 8)), 64)
        Next Counter
    End If
    
    For Counter = 1 To Len(Text) Step 16
        BinaryText = CSM_Conversion.Hex2Bin(Right(Text, 16))
        Text = Left(Text, Len(Text) - 16)
    
        DecryptingBlock = DecryptBinBlock(BinaryText, BinaryKey)
        
        If Text <> "" Then
            PreviousBlock = CSM_Conversion.Hex2Bin(Right(Text, 16))
        Else
            PreviousBlock = EncryptBinBlock(CSM_Conversion.Text2Bin(INIT_VECTOR), BinaryKey)
        End If
        DecryptedBlock = ApplyXOR(PreviousBlock, DecryptingBlock, 64)
        
        DecryptString = CSM_Conversion.Bin2Text(DecryptedBlock) & DecryptString
    Next Counter
    
    TextLength = (CLng(Asc(Mid(DecryptString, 2, 1))) * 256) + CLng(Asc(Mid(DecryptString, 3, 1)))
    If Left(DecryptString, 1) = "T" Then
        DecryptString = Mid(DecryptString, 12, TextLength)
    Else
        DecryptString = Mid(DecryptString, 4, TextLength)
    End If
End Function

Private Function DecryptBinBlock(ByVal BinaryText As String, ByVal BinaryKey As String) As String
    Dim LoRoText As String
    Dim LoText As String
    Dim RoText As String
    
    Dim CoDoKey As String
    Dim CoKey As String
    Dim DoKey As String
    
    Dim ERiText As String
    Dim KiText As String
    Dim XORText As String
    Dim SBoxText As String
    Dim RiKiText As String
    Dim LiRiText As String
    
    Dim IPInverse As String
    
    Dim KeySchedule1_16(1 To 16) As String
    
    Dim IterationNumber As Long
    Dim ShiftPosition As Long
        
    'Apply IP to Text
    LoRoText = ApplyIP(BinaryText)
    LoText = Left(LoRoText, 32)
    RoText = Mid(LoRoText, 33)

    'Apply PC1 to Key
    CoDoKey = ApplyPC1(BinaryKey)
    CoKey = Left(CoDoKey, 28)
    DoKey = Mid(CoDoKey, 29)
    
    'Do 16 Iterations to Apply E Expansion Function and PC2
    For IterationNumber = 1 To 16
        ERiText = ExpansionFunctionE(RoText)
        
        If IterationNumber = 1 Or IterationNumber = 2 Or IterationNumber = 9 Or IterationNumber = 16 Then
            ShiftPosition = 1
        Else
            ShiftPosition = 2
        End If
        
        CoKey = LeftShift(CoKey, ShiftPosition)
        DoKey = LeftShift(DoKey, ShiftPosition)
        CoDoKey = CoKey & DoKey
        
        KiText = ApplyPC2(CoDoKey)
        
        KeySchedule1_16(IterationNumber) = KiText
    Next IterationNumber

    'Do 16 Iterations to Apply E Expansion Function and PC2
    For IterationNumber = 1 To 16
        ERiText = ExpansionFunctionE(RoText)
        
        KiText = KeySchedule1_16(17 - IterationNumber)
        
        XORText = ApplyXOR(ERiText, KiText, 48)
        
        SBoxText = ApplySBoxes(XORText)
        
        RiKiText = ApplyP(SBoxText)
        
        LiRiText = ApplyXOR(LoText, RiKiText, 32)
        
        LoText = RoText
        RoText = LiRiText
    Next IterationNumber
    
    'IP Inverse for L16 and R16
    IPInverse = ApplyIPInverse(RoText & LoText)
    
    DecryptBinBlock = IPInverse
End Function

Private Function ApplyIP(ByVal BinaryText As String) As String
    If Len(BinaryText) > 64 Then
        MsgBox "IP function ..." & vbCr & "This chunk of bits is > 64 !!!", vbExclamation, App.Title
        Exit Function
    End If

    'The initial permutation IP is as follows:
    ApplyIP = Mid(BinaryText, 58, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 50, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 42, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 34, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 26, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 18, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 10, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 2, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 60, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 52, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 44, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 36, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 28, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 20, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 12, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 4, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 62, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 54, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 46, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 38, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 30, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 22, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 14, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 6, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 64, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 56, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 48, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 40, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 32, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 24, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 16, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 8, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 57, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 49, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 41, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 33, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 25, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 17, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 9, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 1, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 59, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 51, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 43, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 35, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 27, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 19, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 11, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 3, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 61, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 53, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 45, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 37, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 29, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 21, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 13, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 5, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 63, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 55, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 47, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 39, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 31, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 23, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 15, 1)
    ApplyIP = ApplyIP & Mid(BinaryText, 7, 1)
End Function

Private Function ApplyPC1(ByVal BinaryKey As String) As String
    If Len(BinaryKey) > 64 Then
        MsgBox "PC1 function ..." & vbCr & "This chunk of bits is > 64 !!!", vbExclamation, App.Title
        Exit Function
    End If

    ApplyPC1 = Mid(BinaryKey, 57, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 49, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 41, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 33, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 25, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 17, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 9, 1)
    
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 1, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 58, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 50, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 42, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 34, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 26, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 18, 1)
    
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 10, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 2, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 59, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 51, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 43, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 35, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 27, 1)
    
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 19, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 11, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 3, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 60, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 52, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 44, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 36, 1)
    
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 63, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 55, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 47, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 39, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 31, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 23, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 15, 1)
    
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 7, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 62, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 54, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 46, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 38, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 30, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 22, 1)
    
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 14, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 6, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 61, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 53, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 45, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 37, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 29, 1)
    
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 21, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 13, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 5, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 28, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 20, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 12, 1)
    ApplyPC1 = ApplyPC1 & Mid(BinaryKey, 4, 1)
End Function

Private Function ExpansionFunctionE(ByVal BinaryValue As String) As String
    If Len(BinaryValue) > 32 Then
        MsgBox "E function ..." & vbCr & "This chunk of bits is > 32 !!!", vbExclamation, App.Title
        Exit Function
    End If
    
    ExpansionFunctionE = Mid(BinaryValue, 32, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 1, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 2, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 3, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 4, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 5, 1)
    
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 4, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 5, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 6, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 7, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 8, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 9, 1)
    
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 8, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 9, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 10, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 11, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 12, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 13, 1)
    
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 12, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 13, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 14, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 15, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 16, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 17, 1)
    
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 16, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 17, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 18, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 19, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 20, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 21, 1)
    
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 20, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 21, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 22, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 23, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 24, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 25, 1)
    
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 24, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 25, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 26, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 27, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 28, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 29, 1)
    
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 28, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 29, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 30, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 31, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 32, 1)
    ExpansionFunctionE = ExpansionFunctionE & Mid(BinaryValue, 1, 1)
End Function

Private Function LeftShift(ByVal BinaryValue As String, ShiftPosition As Long) As String
    LeftShift = Mid(BinaryValue, ShiftPosition + 1) & Left(BinaryValue, ShiftPosition)
End Function

Private Function ApplyPC2(ByVal BinaryKey As String) As String
    If Len(BinaryKey) > 56 Then
        MsgBox "PC2 function ..." & vbCr & "This chunk of bits is > 56 !!!", vbExclamation, App.Title
        Exit Function
    End If
    
    ApplyPC2 = Mid(BinaryKey, 14, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 17, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 11, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 24, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 1, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 5, 1)
    
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 3, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 28, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 15, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 6, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 21, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 10, 1)
    
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 23, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 19, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 12, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 4, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 26, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 8, 1)
    
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 16, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 7, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 27, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 20, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 13, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 2, 1)
    
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 41, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 52, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 31, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 37, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 47, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 55, 1)
    
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 30, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 40, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 51, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 45, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 33, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 48, 1)
    
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 44, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 49, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 39, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 56, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 34, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 53, 1)
    
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 46, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 42, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 50, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 36, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 29, 1)
    ApplyPC2 = ApplyPC2 & Mid(BinaryKey, 32, 1)
End Function

Private Function ApplyXOR(ByVal Value1 As String, ByVal Value2 As String, ByVal Length As Long) As String
    Dim CharPos As Long
    Dim bitValue1 As Byte
    Dim bitValue2 As Byte
    
    For CharPos = 1 To Length
        On Error Resume Next
        bitValue1 = CByte(Mid(Value1, CharPos, 1))
        bitValue2 = CByte(Mid(Value2, CharPos, 1))
        
        ApplyXOR = ApplyXOR & (bitValue1 Xor bitValue2)
    Next CharPos
End Function

Private Function ApplySBoxes(ByVal BinaryValue As String) As String
    Dim IterationNumber As Long
    Dim Bi As String
    Dim b As Byte
    Dim b1_b6 As String
    Dim b2_b5 As String
    Dim bCol As Integer
    Dim bRow As Integer
    Dim SboxNum As Integer
    Dim sbNumber As Integer
    Dim chkSbox As String
    Dim zeros As Long
    Dim z As Long
    
    If Len(BinaryValue) > 48 Then
        MsgBox "SBoxes function ..." & vbCr & "This chunk of bits is > 48 !!!", vbExclamation, App.Title
        Exit Function
    End If
    
    For IterationNumber = 1 To 48 Step 6
        Bi = Mid(BinaryValue, IterationNumber, 6)
    
        'Col.
        'first of 2 bits of Col.
        b = Left(Bi, 1)
        b1_b6 = b
        
        'second of 2 bits of Col.
        b = Right(Bi, 1)
        b1_b6 = b1_b6 & b
    
        'Row
        'four bits of Row
        b2_b5 = Mid(Bi, 2, 4)
    
        'Convert Col abd Row to decimal
        bRow = CSM_Conversion.FromBase(b1_b6, 2)
        bCol = CSM_Conversion.FromBase(b2_b5, 2)
        
        'Now 8 S-boxes
        SboxNum = SboxNum + 1
    
        Select Case SboxNum
            Case 1
                sbNumber = ApplySBox1(bRow, bCol)
            Case 2
                sbNumber = ApplySBox2(bRow, bCol)
            Case 3
                sbNumber = ApplySBox3(bRow, bCol)
            Case 4
                sbNumber = ApplySBox4(bRow, bCol)
            Case 5
                sbNumber = ApplySBox5(bRow, bCol)
            Case 6
                sbNumber = ApplySBox6(bRow, bCol)
            Case 7
                sbNumber = ApplySBox7(bRow, bCol)
            Case 8
                sbNumber = ApplySBox8(bRow, bCol)
        End Select
     
        'Now Convert back to binary
        chkSbox = CSM_Conversion.ToBase(sbNumber, 2)
    
        'we need 4 bits length
        If Len(chkSbox) < 4 Then
            zeros = 4 - Len(chkSbox)
            For z = 1 To zeros
                chkSbox = "0" & chkSbox
            Next z
        End If
    
        'Append to Sbox string
        ApplySBoxes = ApplySBoxes & chkSbox
    Next IterationNumber
End Function

Private Function ApplySBoxN(ByRef aData As Variant, ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    ApplySBoxN = aData(bRow)(bCol)
End Function

Private Function ApplySBox1(ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    Dim Row0 As Variant
    Dim Row1 As Variant
    Dim Row2 As Variant
    Dim Row3 As Variant

    Row0 = Array(14, 4, 13, 1, 2, 15, 11, 8, 3, 10, 6, 12, 5, 9, 0, 7)
    Row1 = Array(0, 15, 7, 4, 14, 2, 13, 1, 10, 6, 12, 11, 9, 5, 3, 8)
    Row2 = Array(4, 1, 14, 8, 13, 6, 2, 11, 15, 12, 9, 7, 3, 10, 5, 0)
    Row3 = Array(15, 12, 8, 2, 4, 9, 1, 7, 5, 11, 3, 14, 10, 0, 6, 13)
        
    ApplySBox1 = ApplySBoxN(Array(Row0, Row1, Row2, Row3), bRow, bCol)
End Function

Private Function ApplySBox2(ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    Dim Row0 As Variant
    Dim Row1 As Variant
    Dim Row2 As Variant
    Dim Row3 As Variant

    Row0 = Array(15, 1, 8, 14, 6, 11, 3, 4, 9, 7, 2, 13, 12, 0, 5, 10)
    Row1 = Array(3, 13, 4, 7, 15, 2, 8, 14, 12, 0, 1, 10, 6, 9, 11, 5)
    Row2 = Array(0, 14, 7, 11, 10, 4, 13, 1, 5, 8, 12, 6, 9, 3, 2, 15)
    Row3 = Array(13, 8, 10, 1, 3, 15, 4, 2, 11, 6, 7, 12, 0, 5, 14, 9)
    
    ApplySBox2 = ApplySBoxN(Array(Row0, Row1, Row2, Row3), bRow, bCol)
End Function

Private Function ApplySBox3(ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    Dim Row0 As Variant
    Dim Row1 As Variant
    Dim Row2 As Variant
    Dim Row3 As Variant

    Row0 = Array(10, 0, 9, 14, 6, 3, 15, 5, 1, 13, 12, 7, 11, 4, 2, 8)
    Row1 = Array(13, 7, 0, 9, 3, 4, 6, 10, 2, 8, 5, 14, 12, 11, 15, 1)
    Row2 = Array(13, 6, 4, 9, 8, 15, 3, 0, 11, 1, 2, 12, 5, 10, 14, 7)
    Row3 = Array(1, 10, 13, 0, 6, 9, 8, 7, 4, 15, 14, 3, 11, 5, 2, 12)
    
    ApplySBox3 = ApplySBoxN(Array(Row0, Row1, Row2, Row3), bRow, bCol)
End Function

Private Function ApplySBox4(ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    Dim Row0 As Variant
    Dim Row1 As Variant
    Dim Row2 As Variant
    Dim Row3 As Variant

    Row0 = Array(7, 13, 14, 3, 0, 6, 9, 10, 1, 2, 8, 5, 11, 12, 4, 15)
    Row1 = Array(13, 8, 11, 5, 6, 15, 0, 3, 4, 7, 2, 12, 1, 10, 14, 9)
    Row2 = Array(10, 6, 9, 0, 12, 11, 7, 13, 15, 1, 3, 14, 5, 2, 8, 4)
    Row3 = Array(3, 15, 0, 6, 10, 1, 13, 8, 9, 4, 5, 11, 12, 7, 2, 14)
    
    ApplySBox4 = ApplySBoxN(Array(Row0, Row1, Row2, Row3), bRow, bCol)
End Function

Private Function ApplySBox5(ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    Dim Row0 As Variant
    Dim Row1 As Variant
    Dim Row2 As Variant
    Dim Row3 As Variant

    Row0 = Array(2, 12, 4, 1, 7, 10, 11, 6, 8, 5, 3, 15, 13, 0, 14, 9)
    Row1 = Array(14, 11, 2, 12, 4, 7, 13, 1, 5, 0, 15, 10, 3, 9, 8, 6)
    Row2 = Array(4, 2, 1, 11, 10, 13, 7, 8, 15, 9, 12, 5, 6, 3, 0, 14)
    Row3 = Array(11, 8, 12, 7, 1, 14, 2, 13, 6, 15, 0, 9, 10, 4, 5, 3)
    
    ApplySBox5 = ApplySBoxN(Array(Row0, Row1, Row2, Row3), bRow, bCol)
End Function

Private Function ApplySBox6(ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    Dim Row0 As Variant
    Dim Row1 As Variant
    Dim Row2 As Variant
    Dim Row3 As Variant

    Row0 = Array(12, 1, 10, 15, 9, 2, 6, 8, 0, 13, 3, 4, 14, 7, 5, 11)
    Row1 = Array(10, 15, 4, 2, 7, 12, 9, 5, 6, 1, 13, 14, 0, 11, 3, 8)
    Row2 = Array(9, 14, 15, 5, 2, 8, 12, 3, 7, 0, 4, 10, 1, 13, 11, 6)
    Row3 = Array(4, 3, 2, 12, 9, 5, 15, 10, 11, 14, 1, 7, 6, 0, 8, 13)
    
    ApplySBox6 = ApplySBoxN(Array(Row0, Row1, Row2, Row3), bRow, bCol)
End Function

Private Function ApplySBox7(ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    Dim Row0 As Variant
    Dim Row1 As Variant
    Dim Row2 As Variant
    Dim Row3 As Variant

    Row0 = Array(4, 11, 2, 14, 15, 0, 8, 13, 3, 12, 9, 7, 5, 10, 6, 1)
    Row1 = Array(13, 0, 11, 7, 4, 9, 1, 10, 14, 3, 5, 12, 2, 15, 8, 6)
    Row2 = Array(1, 4, 11, 13, 12, 3, 7, 14, 10, 15, 6, 8, 0, 5, 9, 2)
    Row3 = Array(6, 11, 13, 8, 1, 4, 10, 7, 9, 5, 0, 15, 14, 2, 3, 12)
    
    ApplySBox7 = ApplySBoxN(Array(Row0, Row1, Row2, Row3), bRow, bCol)
End Function

Private Function ApplySBox8(ByVal bRow As Integer, ByVal bCol As Integer) As Integer
    Dim Row0 As Variant
    Dim Row1 As Variant
    Dim Row2 As Variant
    Dim Row3 As Variant

    Row0 = Array(13, 2, 8, 4, 6, 15, 11, 1, 10, 9, 3, 14, 5, 0, 12, 7)
    Row1 = Array(1, 15, 13, 8, 10, 3, 7, 4, 12, 5, 6, 11, 0, 14, 9, 2)
    Row2 = Array(7, 11, 4, 1, 9, 12, 14, 2, 0, 6, 10, 13, 15, 3, 5, 8)
    Row3 = Array(2, 1, 14, 7, 4, 10, 8, 13, 15, 12, 9, 0, 3, 5, 6, 11)
    
    ApplySBox8 = ApplySBoxN(Array(Row0, Row1, Row2, Row3), bRow, bCol)
End Function

Private Function ApplyP(ByVal BinaryValue As String) As String
    If Len(BinaryValue) > 32 Then
       MsgBox "P function ..." & vbCr & "This chunk of bits is > 32 !!!", vbExclamation, App.Title
    End If
        
    'The P permutation function is specified by the following table:
                    
    '16  7   20  21
    '29  12  28  17
    '1   15  23  26
    '5   18  31  10
    '2   8   24  14
    '32  27  3   9
    '19  13  30  6
    '22  11  4   25

    '16  7   20  21
    ApplyP = Mid(BinaryValue, 16, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 7, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 20, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 21, 1)
    '29  12  28  17
    ApplyP = ApplyP & Mid(BinaryValue, 29, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 12, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 28, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 17, 1)
    '1   15  23  26
    ApplyP = ApplyP & Mid(BinaryValue, 1, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 15, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 23, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 26, 1)
    '5   18  31  10
    ApplyP = ApplyP & Mid(BinaryValue, 5, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 18, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 31, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 10, 1)
    '2   8   24  14
    ApplyP = ApplyP & Mid(BinaryValue, 2, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 8, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 24, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 14, 1)
    '32  27  3   9
    ApplyP = ApplyP & Mid(BinaryValue, 32, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 27, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 3, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 9, 1)
    '19  13  30  6
    ApplyP = ApplyP & Mid(BinaryValue, 19, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 13, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 30, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 6, 1)
    '22  11  4   25
    ApplyP = ApplyP & Mid(BinaryValue, 22, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 11, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 4, 1)
    ApplyP = ApplyP & Mid(BinaryValue, 25, 1)
End Function

Private Function ApplyIPInverse(ByVal Value As String) As String
    If Len(Value) > 64 Then
        MsgBox "IP Reverse function ..." & vbCr & "This chunk of bits is > 64 !!!", vbExclamation, App.Title
    End If
          
    'The IP-1 permutation function is specified by the following table:
                    
    '40  8   48  16  56  24  64  32
    '39  7   47  15  55  23  63  31
    '38  6   46  14  54  22  62  30
    '37  5   45  13  53  21  61  29
    '36  4   44  12  52  20  60  28
    '35  3   43  11  51  19  59  27
    '34  2   42  10  50  18  58  26
    '33  1   41  9   49  17  57  25
    
    
    'The permutation IP-1 inverse is as follows:
        
    '40  8   48  16  56  24  64  32
    ApplyIPInverse = Mid(Value, 40, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 8, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 48, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 16, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 56, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 24, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 64, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 32, 1)
    '39  7   47  15  55  23  63  31
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 39, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 7, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 47, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 15, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 55, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 23, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 63, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 31, 1)
    '38  6   46  14  54  22  62  30
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 38, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 6, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 46, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 14, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 54, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 22, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 62, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 30, 1)
    '37  5   45  13  53  21  61  29
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 37, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 5, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 45, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 13, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 53, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 21, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 61, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 29, 1)
    '36  4   44  12  52  20  60  28
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 36, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 4, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 44, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 12, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 52, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 20, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 60, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 28, 1)
    '35  3   43  11  51  19  59  27
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 35, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 3, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 43, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 11, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 51, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 19, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 59, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 27, 1)
    '34  2   42  10  50  18  58  26
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 34, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 2, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 42, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 10, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 50, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 18, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 58, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 26, 1)
    '33  1   41  9   49  17  57  25
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 33, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 1, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 41, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 9, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 49, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 17, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 57, 1)
    ApplyIPInverse = ApplyIPInverse & Mid(Value, 25, 1)
End Function


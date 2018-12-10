Attribute VB_Name = "CSM_String"
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long

Public Function ReplaceQuote(ByVal Expression As String) As String
    ReplaceQuote = Replace(Expression, "'", "''")
End Function

Public Function GetSubString(ByVal MainString As String, ByVal SubStringPosition As Integer, ByVal SubStringSeparator As String) As String
    Dim aArray() As String
    
    aArray = Split(MainString, SubStringSeparator)
    If SubStringPosition <= UBound(aArray) + 1 Then
        GetSubString = aArray(SubStringPosition - 1)
    End If
End Function

Public Function GetSubStringAsByte(ByVal MainString As String, ByVal SubStringPosition As Integer, ByVal SubStringSeparator As String) As Byte
    Dim TempValue As String
    
    TempValue = GetSubString(MainString, SubStringPosition, SubStringSeparator)
    If IsNumeric(TempValue) Then
        If Val(TempValue) < 256 Then
            GetSubStringAsByte = CByte(TempValue)
        Else
            GetSubStringAsByte = 0
        End If
    Else
        GetSubStringAsByte = 0
    End If
End Function

Public Function RemoveFirstSubString(ByVal MainString As String, ByVal SubStringSeparator As String) As String
    Dim aArray() As String
    
    If Len(MainString) > 0 Then
        aArray = Split(MainString, SubStringSeparator)
        RemoveFirstSubString = Mid(MainString, Len(aArray(LBound(aArray))) + 2)
    End If
End Function

Public Function RemoveLastSubString(ByVal MainString As String, ByVal SubStringSeparator As String) As String
    Dim aArray() As String
    
    If Len(MainString) > 0 Then
        aArray = Split(MainString, SubStringSeparator)
        If UBound(aArray) > 0 Then
            RemoveLastSubString = Left(MainString, Len(MainString) - Len(aArray(UBound(aArray))) - 1)
        Else
            RemoveLastSubString = MainString
        End If
    End If
End Function

Public Function CountSubString(ByVal MainString As String, ByVal SubStringSeparator As String) As String
    Dim aArray() As String
    
    aArray = Split(MainString, SubStringSeparator)
    CountSubString = UBound(aArray) + 1
End Function

Public Function GetSubStringValue(ByVal MainString As String, ByVal SubStringSeparator As String, ByVal ValueSeparator As String, ByVal ValueName As String) As String
    Dim aArray() As String
    Dim ValueIndex As Integer
    Dim CurrentValueName As String
    
    aArray = Split(MainString, SubStringSeparator)
    For ValueIndex = 0 To UBound(aArray)
        CurrentValueName = GetSubString(aArray(ValueIndex), 1, ValueSeparator)
        If CurrentValueName = ValueName Then
            GetSubStringValue = GetSubString(aArray(ValueIndex), 2, ValueSeparator)
            Exit For
        End If
    Next ValueIndex
End Function

Public Function GetBooleanString(ByVal Value As Boolean) As String
    GetBooleanString = IIf(Value, "Sí", "No")
End Function

Public Function GetBooleanValueFromString(ByVal Value As String) As Boolean
    GetBooleanValueFromString = (Value = "Sí")
End Function

Public Function CleanInvalidCharsByAllowed(ByVal Value As String, ByVal AllowedChars As String) As String
    Dim CharIndex As Long
    
    For CharIndex = 1 To Len(Value)
        If InStr(1, AllowedChars, Mid(Value, CharIndex, 1)) > 0 Then
            CleanInvalidCharsByAllowed = CleanInvalidCharsByAllowed & Mid(Value, CharIndex, 1)
        End If
    Next CharIndex
End Function

Public Function CleanInvalidCharsByNotAllowed(ByVal Value As String, ByVal NotAllowedChars As String) As String
    Dim CharIndex As Long
    
    For CharIndex = 1 To Len(Value)
        If InStr(1, NotAllowedChars, Mid(Value, CharIndex, 1)) = 0 Then
            CleanInvalidCharsByNotAllowed = CleanInvalidCharsByNotAllowed & Mid(Value, CharIndex, 1)
        End If
    Next CharIndex
End Function

Public Function CleanNotNumericChars(ByVal Value As String) As String
    CleanNotNumericChars = CleanInvalidCharsByAllowed(Value, "0123456789")
End Function

Public Function CleanInvalidSpaces(ByVal Value As String) As String
    CleanInvalidSpaces = Trim(Value)
    
    Do While InStr(1, CleanInvalidSpaces, "  ") > 0
        CleanInvalidSpaces = Replace(CleanInvalidSpaces, "  ", " ")
    Loop
End Function

Public Function CleanNullChars(ByVal Value As String) As String
    CleanNullChars = Value
    
    Do While InStr(1, CleanNullChars, vbNullChar) > 0
        CleanNullChars = Replace(CleanNullChars, vbNullChar, "")
    Loop
End Function

Public Function ConvertCurrencyToVBNumber(ByVal Value As Currency) As String
    ConvertCurrencyToVBNumber = Replace(Format(Value), pRegionalSettings.CurrencyDecimalSymbol, ".")
End Function

Public Function ConvertDoubleToVBNumber(ByVal Value As Double) As String
    ConvertDoubleToVBNumber = Replace(Format(Value), pRegionalSettings.NumberDecimalSymbol, ".")
End Function

Public Function FormatDoubleToString_NoGrouping_DotAsDecimal(ByVal Value As Double, ByVal FormatExpression As String) As String
    FormatDoubleToString_NoGrouping_DotAsDecimal = Replace(Replace(Format(Value, FormatExpression), pRegionalSettings.NumberDigitGroupingSymbol, ""), pRegionalSettings.NumberDecimalSymbol, ".")
End Function

Public Function FormatDoubleToString_NoGrouping_CommaAsDecimal(ByVal Value As Double, ByVal FormatExpression As String) As String
    FormatDoubleToString_NoGrouping_CommaAsDecimal = Replace(Replace(Format(Value, FormatExpression), pRegionalSettings.NumberDigitGroupingSymbol, ""), pRegionalSettings.NumberDecimalSymbol, ",")
End Function

Public Function GetStringExtentInPixels(ByVal hdc As Long, ByVal Value As String) As Long
    Dim TextSize As POINTAPI
    Dim lngResult As Long
    
    lngResult = GetTextExtentPoint32(hdc, Value, Len(Value), TextSize)
    If lngResult <> 0 Then
        GetStringExtentInPixels = TextSize.x
    End If
End Function

Public Function RemoveCharsAfterNull(ByVal Value As String) As String
    Dim FirstNullPosition As Long

    FirstNullPosition = InStr(1, Value, vbNullChar)
    If FirstNullPosition > 0 Then
        RemoveCharsAfterNull = Left(Value, FirstNullPosition - 1)
    Else
        RemoveCharsAfterNull = Value
    End If
End Function

Public Function PadString(ByVal Value As String, ByVal FillerComplete As String, ByVal FillLeft As Boolean, ByVal FillRight As Boolean, Optional ByVal TrimString As Boolean = True) As String
    If Len(Value) >= Len(FillerComplete) Then
        If TrimString Then
            If FillLeft And FillRight Then
                'CENTER
                PadString = Mid(Value, ((Len(FillerComplete) - Len(Value)) \ 2), ((Len(FillerComplete) - Len(Value)) \ 2) + IIf((Len(FillerComplete) - Len(Value)) Mod 2 > 0, 1, 0))
            ElseIf FillLeft Then
                'LEFT
                PadString = Right(Value, Len(FillerComplete))
            ElseIf FillRight Then
                'RIGHT
                PadString = Left(Value, Len(FillerComplete))
            Else
                'NONE
                PadString = Value
            End If
        Else
            'DON'T TRIM
            PadString = Value
        End If
    Else
        If FillLeft And FillRight Then
            'CENTER
            PadString = Left(FillerComplete, (Len(FillerComplete) - Len(Value)) \ 2) & Value & Right(FillerComplete, ((Len(FillerComplete) - Len(Value)) \ 2) + IIf((Len(FillerComplete) - Len(Value)) Mod 2 > 0, 1, 0))
        ElseIf FillLeft Then
            'LEFT
            PadString = Left(FillerComplete, Len(FillerComplete) - Len(Value)) & Value
        ElseIf FillRight Then
            'RIGHT
            PadString = Value & Right(FillerComplete, Len(FillerComplete) - Len(Value))
        Else
            'NONE
            PadString = Value
        End If
    End If
End Function

Public Function PadStringLeft(ByVal Value As String, ByVal FillerChar As String, ByVal CharCount As Integer) As String
    PadStringLeft = PadString(Value, String(CharCount, FillerChar), True, False)
End Function

Public Function PadStringRight(ByVal Value As String, ByVal FillerChar As String, ByVal CharCount As Integer) As String
    PadStringRight = PadString(Value, String(CharCount, FillerChar), False, True)
End Function

Public Function PadStringCenter(ByVal Value As String, ByVal FillerChar As String, ByVal CharCount As Integer) As String
    PadStringCenter = PadString(Value, String(CharCount, FillerChar), True, True)
End Function

Public Function ConvertCurrencyToWords(ByVal Importe As Currency) As String
    If Importe < 0 Then
        Importe = Abs(Importe)
        ConvertCurrencyToWords = "Menos "
    Else
        ConvertCurrencyToWords = ""
    End If
    ConvertCurrencyToWords = ConvertCurrencyToWords & ConvertIntegerToWords(Int(Importe)) & " con " & Format(((Importe - Int(Importe)) * 100), "00") & "/100"
End Function

Public Function ConvertIntegerToWords(ByVal Number As Long) As String
    Select Case Number
        Case 0
            ConvertIntegerToWords = ""
        Case Is < 0
            ConvertIntegerToWords = "menor que cero"
        Case 1
            ConvertIntegerToWords = "uno"
        Case 2
            ConvertIntegerToWords = "dos"
        Case 3
            ConvertIntegerToWords = "tres"
        Case 4
            ConvertIntegerToWords = "cuatro"
        Case 5
            ConvertIntegerToWords = "cinco"
        Case 6
            ConvertIntegerToWords = "seis"
        Case 7
            ConvertIntegerToWords = "siete"
        Case 8
            ConvertIntegerToWords = "ocho"
        Case 9
            ConvertIntegerToWords = "nueve"
        Case 10
            ConvertIntegerToWords = "diez"
        Case 11
            ConvertIntegerToWords = "once"
        Case 12
            ConvertIntegerToWords = "doce"
        Case 13
            ConvertIntegerToWords = "trece"
        Case 14
            ConvertIntegerToWords = "catorce"
        Case 15
            ConvertIntegerToWords = "quince"
        Case 16
            ConvertIntegerToWords = "dieciseis"
        Case 17
            ConvertIntegerToWords = "diecisiete"
        Case 18
            ConvertIntegerToWords = "dieciocho"
        Case 19
            ConvertIntegerToWords = "diecinueve"
        Case 20
            ConvertIntegerToWords = "veinte"
        Case 21 To 29
            ConvertIntegerToWords = "veinti" + ConvertIntegerToWords(Number - 20)
        Case 30
            ConvertIntegerToWords = "treinta"
        Case 31 To 39
            ConvertIntegerToWords = "treinta y " + ConvertIntegerToWords(Number - 30)
        Case 40
            ConvertIntegerToWords = "cuarenta"
        Case 41 To 49
            ConvertIntegerToWords = "cuarenta y " + ConvertIntegerToWords(Number - 40)
        Case 50
            ConvertIntegerToWords = "cincuenta"
        Case 51 To 59
            ConvertIntegerToWords = "cincuenta y " + ConvertIntegerToWords(Number - 50)
        Case 60
            ConvertIntegerToWords = "sesenta"
        Case 61 To 69
            ConvertIntegerToWords = "sesenta y " + ConvertIntegerToWords(Number - 60)
        Case 70
            ConvertIntegerToWords = "setenta"
        Case 71 To 79
            ConvertIntegerToWords = "setenta y " + ConvertIntegerToWords(Number - 70)
        Case 80
            ConvertIntegerToWords = "ochenta"
        Case 81 To 89
            ConvertIntegerToWords = "ochenta y " + ConvertIntegerToWords(Number - 80)
        Case 90
            ConvertIntegerToWords = "noventa"
        Case 91 To 99
            ConvertIntegerToWords = "noventa y " + ConvertIntegerToWords(Number - 90)
        Case 100
            ConvertIntegerToWords = "cien"
        Case 101 To 199
            ConvertIntegerToWords = "ciento " + ConvertIntegerToWords(Number - 100)
        Case 200
            ConvertIntegerToWords = "doscientos"
        Case 201 To 299
            ConvertIntegerToWords = "doscientos " + ConvertIntegerToWords(Number - 200)
        Case 300
            ConvertIntegerToWords = "trescientos"
        Case 301 To 399
            ConvertIntegerToWords = "trescientos " + ConvertIntegerToWords(Number - 300)
        Case 400
            ConvertIntegerToWords = "cuatrocientos"
        Case 401 To 499
            ConvertIntegerToWords = "cuatrocientos " + ConvertIntegerToWords(Number - 400)
        Case 500
            ConvertIntegerToWords = "quinientos"
        Case 501 To 599
            ConvertIntegerToWords = "quinientos " + ConvertIntegerToWords(Number - 500)
        Case 600
            ConvertIntegerToWords = "seiscientos"
        Case 601 To 699
            ConvertIntegerToWords = "seiscientos " + ConvertIntegerToWords(Number - 600)
        Case 700
            ConvertIntegerToWords = "setecientos"
        Case 701 To 799
            ConvertIntegerToWords = "setecientos " + ConvertIntegerToWords(Number - 700)
        Case 800
            ConvertIntegerToWords = "ochocientos"
        Case 801 To 899
            ConvertIntegerToWords = "ochocientos " + ConvertIntegerToWords(Number - 800)
        Case 900
            ConvertIntegerToWords = "novecientos"
        Case 901 To 999
            ConvertIntegerToWords = "novecientos " + ConvertIntegerToWords(Number - 900)
        Case 1000
            ConvertIntegerToWords = "un mil"
        Case 1001 To 1999
            ConvertIntegerToWords = "un mil " + ConvertIntegerToWords(Number Mod 1000)
        Case 2000 To 1000000
            ConvertIntegerToWords = ConvertIntegerToWords(Int(Number / 1000)) + " mil " + ConvertIntegerToWords(Number Mod 1000)
        Case 1000000 To 2000000
            ConvertIntegerToWords = "un millon " + ConvertIntegerToWords(Number - 1000000)
        Case 2000000 To 1E+20
            ConvertIntegerToWords = ConvertIntegerToWords(Int(Number / 1000000)) + " millones " + ConvertIntegerToWords(Number Mod 1000000)
    End Select
End Function

Public Function ConvertDelimitedStringToCollection(ByVal DelimitedString As String, ByVal SubStringSeparator As String) As Collection
    Set ConvertDelimitedStringToCollection = ConvertArrayToCollection(Split(DelimitedString, SubStringSeparator))
End Function

Public Function ConvertArrayToCollection(ByVal aArray As Variant) As Collection
    Dim Index As Integer
    
    If IsArray(aArray) Then
        Set ConvertArrayToCollection = New Collection
        If UBound(aArray) > -1 Then
            For Index = 0 To UBound(aArray)
                ConvertArrayToCollection.Add aArray(Index)
            Next Index
        End If
    End If
End Function

Public Function FormatStringToSQL(ByVal Value As String, Optional ConvertEmptyToNull As Boolean = False) As String
    If Trim(Value) = "" And ConvertEmptyToNull Then
        FormatStringToSQL = "NULL"
    Else
        FormatStringToSQL = "'" & ReplaceQuote(Value) & "'"
    End If
End Function

Public Function FormatIntegerToSQL(ByVal Value As Long, Optional ConvertZeroToNull As Boolean = False) As String
    If Value = 0 And ConvertZeroToNull Then
        FormatIntegerToSQL = "NULL"
    Else
        FormatIntegerToSQL = Value
    End If
End Function

Public Function FormatDecimalToSQL(ByVal Value As Single, Optional ConvertZeroToNull As Boolean = False) As String
    If Value = 0 And ConvertZeroToNull Then
        FormatDecimalToSQL = "NULL"
    Else
        FormatDecimalToSQL = Replace(Value, pRegionalSettings.NumberDecimalSymbol, ".")
    End If
End Function

Public Function FormatCurrencyToSQL(ByVal Value As Currency, Optional ConvertZeroToNull As Boolean = False) As String
    If Value = 0 And ConvertZeroToNull Then
        FormatCurrencyToSQL = "NULL"
    Else
        FormatCurrencyToSQL = Replace(Value, pRegionalSettings.NumberDecimalSymbol, ".")
    End If
End Function

Public Function FormatDateTimeToSQL(ByVal Value As Date, Optional ConvertEmptyToNull As Boolean = False) As String
    If Value = DATE_TIME_FIELD_NULL_VALUE And ConvertEmptyToNull Then
        FormatDateTimeToSQL = "NULL"
    Else
        FormatDateTimeToSQL = "'" & Format(Value, "yyyy/mm/dd hh:nn") & "'"
    End If
End Function

Public Function FormatBooleanToSQL(ByVal Value As Boolean) As String
    If Value Then
        FormatBooleanToSQL = "1"
    Else
        FormatBooleanToSQL = "2"
    End If
End Function

Public Function GetOLEColorFromString(ByVal CommaSeparatedValue As String) As OLE_COLOR
    Dim ColorComponentRed As Byte
    Dim ColorComponentGreen As Byte
    Dim ColorComponentBlue As Byte
    
    ColorComponentRed = GetSubStringAsByte(CommaSeparatedValue, 1, ",")
    ColorComponentGreen = GetSubStringAsByte(CommaSeparatedValue, 2, ",")
    ColorComponentBlue = GetSubStringAsByte(CommaSeparatedValue, 3, ",")
    
    GetOLEColorFromString = RGB(ColorComponentRed, ColorComponentGreen, ColorComponentBlue)
End Function

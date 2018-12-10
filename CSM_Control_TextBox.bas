Attribute VB_Name = "CSM_Control_TextBox"
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////
'
'DATA TYPES:
'===========
'INTEGER
'   EMPTY | NOTEMPTY
'       ZERO | NOTZERO
'           POSITIVE | NEGATIVE
'               OPTIONAL KEYS ALLOWED
'
'DECIMAL
'   EMPTY | NOTEMPTY
'       ZERO | NOTZERO
'           POSITIVE | NEGATIVE
'               MASK (SAMPLE: 99.99)
'                   OPTIONAL KEYS ALLOWED
'
'CURRENCY
'   EMPTY | NOTEMPTY
'       ZERO | NOTZERO
'           POSITIVE | NEGATIVE
'               OPTIONAL KEYS ALLOWED
'
'STRING
'   EMPTY | NOTEMPTY
'       NONE | UPPER | LOWER | CAPITAL | NUMBERS
'           MAXLENGHT
'
'///////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////

Public GoToNextControlOnReturnKeyPressed As Boolean

Public Sub SelAllText(ByRef TheControl As TextBox)
    On Error Resume Next
    
    TheControl.SelStart = 0
    TheControl.SelLength = Len(TheControl.Text)
End Sub

Public Function CheckKeyDown(ByRef TheControl As Control, ByVal KeyCode As Integer) As Boolean
    If TypeOf TheControl Is TextBox Then
        If Not TheControl.Locked Then
            If TheControl.Tag <> "" Then
                If KeyCode = vbKeyDecimal Then
                    CheckKeyDown = True
                End If
            End If
        End If
    End If
End Function

Public Sub CheckKeyPress(ByRef TheControl As Control, ByRef KeyAscii As Integer, ByRef KeyDecimal As Boolean)
    Dim DataTypeString As String
    Dim AllowEmpty As Boolean
    Dim AllowZero As Boolean
    Dim AllowNegative As Boolean
    Dim MaskString As String
    Dim DecimalPlaces As Byte
    Dim Casing As String
    Dim AllowedChars As String
    
    If TypeOf TheControl Is TextBox Then
        If Not TheControl.Locked Then
            If TheControl.Tag <> "" Then
                Call ParseConfig(TheControl.Tag, DataTypeString, AllowEmpty, AllowZero, AllowNegative, MaskString, DecimalPlaces, Casing, AllowedChars)
                
                Select Case DataTypeString
                    Case "INTEGER"
                        Call CheckKeyPress_NumberInteger(KeyAscii, AllowNegative, AllowedChars)
                    Case "DECIMAL"
                        Call CheckKeyPress_NumberDecimal(KeyAscii, KeyDecimal, TheControl.Text, AllowNegative, AllowedChars)
                    Case "CURRENCY"
                        Call CheckKeyPress_Currency(KeyAscii, KeyDecimal, TheControl.Text, AllowNegative, AllowedChars)
                    Case "STRING"
                        Call CheckKeyPress_String(KeyAscii, Casing)
                End Select
            End If
            If GoToNextControlOnReturnKeyPressed Then
                If KeyAscii = vbKeyReturn Then
                    KeyAscii = 0
                    SendKeys "{TAB}"
                End If
            End If
        End If
    End If
    KeyDecimal = False
End Sub

Private Sub CheckKeyPress_Currency(ByRef KeyAscii As Integer, ByRef KeyDecimal As Boolean, ByVal TextCurrent As String, ByVal AllowNegative As Boolean, Optional ByVal AllowedChars As String = "")
    If AllowedChars = "" Then
        If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) And (KeyAscii <> 45 Or Not AllowNegative) Then
            KeyAscii = 0
        End If
    Else
        If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.CurrencyDecimalSymbol) And KeyAscii <> Asc(pRegionalSettings.CurrencyCurrencySymbol) And (KeyAscii <> 45 Or Not AllowNegative) And InStr(1, AllowedChars, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol)
        KeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyDecimalSymbol) Then
        If InStr(1, TextCurrent, pRegionalSettings.CurrencyDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = Asc(pRegionalSettings.CurrencyCurrencySymbol) Then
        If InStr(1, TextCurrent, pRegionalSettings.CurrencyCurrencySymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub CheckKeyPress_NumberInteger(ByRef KeyAscii As Integer, ByVal AllowNegative As Boolean, Optional ByVal AllowedChars As String = "")
    If AllowedChars = "" Then
        If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And (KeyAscii <> 45 Or Not AllowNegative) Then
            KeyAscii = 0
        End If
    Else
        If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And (KeyAscii <> 45 Or Not AllowNegative) And InStr(1, AllowedChars, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub CheckKeyPress_NumberDecimal(ByRef KeyAscii As Integer, ByRef KeyDecimal As Boolean, ByVal TextCurrent As String, ByVal AllowNegative As Boolean, Optional ByVal AllowedChars As String = "")
    If AllowedChars = "" Then
        If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.NumberDecimalSymbol) And (KeyAscii <> 45 Or Not AllowNegative) Then
            KeyAscii = 0
        End If
    Else
        If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> Asc(pRegionalSettings.NumberDecimalSymbol) And (KeyAscii <> 45 Or Not AllowNegative) And InStr(1, AllowedChars, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    If KeyDecimal Then
        KeyAscii = Asc(pRegionalSettings.NumberDecimalSymbol)
        KeyDecimal = False
    End If
    If KeyAscii = Asc(pRegionalSettings.NumberDecimalSymbol) Then
        If InStr(1, TextCurrent, pRegionalSettings.NumberDecimalSymbol) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub CheckKeyPress_String(ByRef KeyAscii As Integer, ByVal Casing As String)
    Select Case Casing
        Case "NONE"
        Case "UPPER"
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case "LOWER"
            KeyAscii = Asc(LCase(Chr(KeyAscii)))
        Case "CAPITAL"
        Case "NUMBERS"
            If ((KeyAscii > 32 And KeyAscii < 48) Or KeyAscii > 57) And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
    End Select
End Sub

Public Sub PrepareAll(ByRef ContainerForm As Form)
    Dim TextBox As Control
    
    For Each TextBox In ContainerForm.Controls
        If TypeOf TextBox Is TextBox Then
'            If Not TextBox.Locked Then
                If TextBox.Tag <> "" Then
                    Call Prepare_ByTag(TextBox)
                    Call FormatValue_ByTag(TextBox)
                End If
'            End If
        End If
    Next TextBox
End Sub

Public Sub Prepare_ByTag(ByRef TextBox As TextBox)
    Dim MaxNumber As String
    Dim IntegerCount As Byte
    Dim DecimalCount As Byte
    
    Dim DataTypeString As String
    Dim AllowEmpty As Boolean
    Dim AllowZero As Boolean
    Dim AllowNegative As Boolean
    Dim MaskString As String
    Dim DecimalPlaces As Byte
    Dim Casing As String
    Dim MaxLenght As Integer
    Dim AllowedChars As String
    
    Call ParseConfig(TextBox.Tag, DataTypeString, AllowEmpty, AllowZero, AllowNegative, MaskString, DecimalPlaces, Casing, AllowedChars)
        
    Select Case CSM_String.GetSubString(TextBox.Tag, 1, "|")
        Case "INTEGER"
            MaxNumber = CSM_String.GetSubString(TextBox.Tag, 5, "|")
            TextBox.Alignment = vbRightJustify
            'TextBox.MaxLength = Len(MaxNumber)
        Case "DECIMAL"
            MaxNumber = CSM_String.GetSubString(TextBox.Tag, 5, "|")
            TextBox.Alignment = vbRightJustify
            DecimalCount = Len(CSM_String.GetSubString(MaxNumber, 2, "."))
            IntegerCount = Len(MaxNumber) - DecimalCount
            If InStr(1, MaxNumber, ".") > 0 Then
                IntegerCount = IntegerCount - 1
            End If
            IntegerCount = IntegerCount + ((IntegerCount - 1) \ pRegionalSettings.NumberNumberOfDigitsInGroup)
            TextBox.MaxLength = IntegerCount
            If DecimalCount > 0 Then
                TextBox.MaxLength = TextBox.MaxLength + 1 + DecimalCount
            End If
        Case "CURRENCY"
            TextBox.Alignment = vbRightJustify
        Case "STRING"
            MaxLenght = Val(CSM_String.GetSubString(TextBox.Tag, 4, "|"))
            TextBox.MaxLength = MaxLenght
    End Select
End Sub

Public Sub FormatAll(ByRef ContainerForm As Form)
    Dim TextBox As Control
    
    For Each TextBox In ContainerForm.Controls
        If TypeOf TextBox Is TextBox Then
            If Not TextBox.Locked Then
                If TextBox.Tag <> "" Then
                    Call FormatValue_ByTag(TextBox)
                End If
            End If
        End If
    Next TextBox
End Sub

Public Sub FormatValue_ByTag(ByRef TextBox As TextBox)
    Dim DataTypeString As String
    Dim AllowEmpty As Boolean
    Dim AllowZero As Boolean
    Dim AllowNegative As Boolean
    Dim MaskString As String
    Dim DecimalPlaces As Byte
    Dim Casing As String
    Dim AllowedChars As String
    
    Call ParseConfig(TextBox.Tag, DataTypeString, AllowEmpty, AllowZero, AllowNegative, MaskString, DecimalPlaces, Casing, AllowedChars)
    Select Case DataTypeString
        Case "INTEGER"
            TextBox.Text = FormatValue_NumberInteger(TextBox.Text, AllowEmpty, AllowNegative, AllowZero)
        Case "DECIMAL"
            TextBox.Text = FormatValue_NumberDecimal(TextBox.Text, AllowEmpty, AllowNegative, AllowZero, DecimalPlaces)
        Case "CURRENCY"
            TextBox.Text = FormatValue_Currency(TextBox.Text, AllowEmpty, AllowNegative, AllowZero)
        Case "STRING"
            TextBox.Text = FormatValue_String(TextBox.Text, Casing)
    End Select
End Sub

Public Function FormatValue_Currency(ByVal TextBoxValue As String, ByVal AllowEmpty As Boolean, ByVal AllowNegative As Boolean, ByVal AllowZero As Boolean) As String
ReDoIt:
    TextBoxValue = Trim(TextBoxValue)
    If TextBoxValue = "" Then
        If (Not AllowEmpty) And AllowZero Then
            TextBoxValue = 0
        End If
    Else
        If Not IsNumeric(TextBoxValue) Then
            TextBoxValue = CSM_String.CleanInvalidCharsByAllowed(TextBoxValue, "0123456789 " & pRegionalSettings.CurrencyCurrencySymbol & pRegionalSettings.CurrencyDecimalSymbol & pRegionalSettings.CurrencyDigitGroupingSymbol & IIf(AllowNegative, "-", ""))
            If Not IsNumeric(TextBoxValue) Then
                TextBoxValue = ""
                Exit Function
            End If
            GoTo ReDoIt
        Else
            Select Case CCur(TextBoxValue)
                Case Is < 0
                    If Not AllowNegative Then
                        TextBoxValue = Abs(CCur(TextBoxValue))
                        GoTo ReDoIt
                    End If
                Case 0
                    If Not AllowZero Then
                        TextBoxValue = ""
                    End If
            End Select
        End If
    End If
    If TextBoxValue <> "" Then
        FormatValue_Currency = Format(CCur(TextBoxValue), "Currency")
    End If
End Function

Public Function FormatValue_NumberInteger(ByVal TextBoxValue As String, ByVal AllowEmpty As Boolean, ByVal AllowNegative As Boolean, ByVal AllowZero As Boolean) As String
ReDoIt:
    TextBoxValue = Trim(TextBoxValue)
    If TextBoxValue = "" Then
        If Not AllowEmpty Then
            TextBoxValue = 0
        End If
    Else
        If Not IsNumeric(TextBoxValue) Then
            TextBoxValue = CSM_String.CleanInvalidCharsByAllowed(TextBoxValue, "0123456789 " & pRegionalSettings.NumberDigitGroupingSymbol & IIf(AllowNegative, "-", ""))
            GoTo ReDoIt
        Else
            If Val(TextBoxValue) > CSM_Constant.DATATYPE_LONG_VALUE_MAX Then
                'Se ingresó un valor muy grande
                Exit Function
            Else
                Select Case CLng(TextBoxValue)
                    Case Is < 0
                        If Not AllowNegative Then
                            TextBoxValue = Abs(CLng(TextBoxValue))
                            GoTo ReDoIt
                        End If
                    Case 0
                        If Not AllowZero Then
                            TextBoxValue = ""
                        End If
                End Select
            End If
        End If
    End If
    If TextBoxValue <> "" Then
        FormatValue_NumberInteger = Format(CLng(TextBoxValue), "#,##0")
    End If
End Function


Public Function FormatValue_NumberDecimal(ByVal TextBoxValue As String, ByVal AllowEmpty As Boolean, ByVal AllowNegative As Boolean, ByVal AllowZero As Boolean, ByVal DecimalPlaces As Byte) As String
    Dim AlreadyRedoit As Boolean
    
ReDoIt:
    TextBoxValue = Trim(TextBoxValue)
    If TextBoxValue = "" Then
        If Not AllowEmpty Then
            TextBoxValue = 0
        End If
    Else
        If Not IsNumeric(TextBoxValue) Then
            TextBoxValue = CSM_String.CleanInvalidCharsByAllowed(TextBoxValue, "0123456789 " & pRegionalSettings.NumberDecimalSymbol & pRegionalSettings.NumberDigitGroupingSymbol & IIf(AllowNegative, "-", ""))
            If AlreadyRedoit Then
                TextBoxValue = ""
            End If
            AlreadyRedoit = True
            GoTo ReDoIt
        Else
            Select Case CDbl(TextBoxValue)
                Case Is < 0
                    If Not AllowNegative Then
                        TextBoxValue = Abs(CDbl(TextBoxValue))
                        GoTo ReDoIt
                    End If
                Case 0
                    If Not AllowZero Then
                        TextBoxValue = ""
                    End If
            End Select
        End If
    End If
    If TextBoxValue <> "" Then
        If DecimalPlaces > 0 Then
            If CDbl(TextBoxValue) - CLng(TextBoxValue) = 0 Then
                FormatValue_NumberDecimal = Format(CDbl(TextBoxValue), "#,##0")
            Else
                FormatValue_NumberDecimal = Format(CDbl(TextBoxValue), "#,##0." & String(DecimalPlaces, "#"))
            End If
        Else
            FormatValue_NumberDecimal = Format(CDbl(TextBoxValue), "#,##0")
        End If
    End If
End Function

Public Function FormatValue_String(ByVal TextBoxValue As String, ByVal Casing As String) As String
    Dim CharIndex As Integer
    Dim UpCaseNextChar As Boolean
    
    Select Case Casing
        Case "", "NONE"
            FormatValue_String = TextBoxValue
        Case "UPPER"
            FormatValue_String = UCase(TextBoxValue)
        Case "LOWER"
            FormatValue_String = LCase(TextBoxValue)
        Case "CAPITAL"
            UpCaseNextChar = True
            For CharIndex = 1 To Len(TextBoxValue)
                If InStr(1, " ,.;:""'()¡!+{}/<>¿?", Mid(TextBoxValue, CharIndex, 1)) <> 0 Then
                    UpCaseNextChar = True
                    FormatValue_String = FormatValue_String & LCase(Mid(TextBoxValue, CharIndex, 1))
                Else
                    If UpCaseNextChar Then
                        UpCaseNextChar = False
                        FormatValue_String = FormatValue_String & UCase(Mid(TextBoxValue, CharIndex, 1))
                    Else
                        FormatValue_String = FormatValue_String & LCase(Mid(TextBoxValue, CharIndex, 1))
                    End If
                End If
            Next CharIndex
        Case "NUMBERS"
            FormatValue_String = CSM_String.CleanNotNumericChars(TextBoxValue)
    End Select
End Function

Private Sub ParseConfig(ByVal ConfigString As String, ByRef DataTypeString As String, ByRef AllowEmpty As Boolean, ByRef AllowZero As Boolean, ByRef AllowNegative As Boolean, ByRef MaskString As String, ByRef DecimalPlaces As Byte, ByRef Casing As String, ByRef AllowedChars As String)
    DataTypeString = CSM_String.GetSubString(ConfigString, 1, "|")
    AllowEmpty = (CSM_String.GetSubString(ConfigString, 2, "|") = "EMPTY")
    Select Case DataTypeString
        Case "INTEGER"
            AllowZero = (CSM_String.GetSubString(ConfigString, 3, "|") = "ZERO")
            AllowNegative = (CSM_String.GetSubString(ConfigString, 4, "|") = "NEGATIVE")
            AllowedChars = CSM_String.GetSubString(ConfigString, 5, "|")
        Case "DECIMAL"
            AllowZero = (CSM_String.GetSubString(ConfigString, 3, "|") = "ZERO")
            AllowNegative = (CSM_String.GetSubString(ConfigString, 4, "|") = "NEGATIVE")
            MaskString = CSM_String.GetSubString(ConfigString, 5, "|")
            DecimalPlaces = Len(CSM_String.GetSubString(MaskString, 2, "."))
            AllowedChars = CSM_String.GetSubString(ConfigString, 6, "|")
        Case "CURRENCY"
            AllowZero = (CSM_String.GetSubString(ConfigString, 3, "|") = "ZERO")
            AllowNegative = (CSM_String.GetSubString(ConfigString, 4, "|") = "NEGATIVE")
            AllowedChars = CSM_String.GetSubString(ConfigString, 5, "|")
        Case "STRING"
            Casing = CSM_String.GetSubString(ConfigString, 3, "|")
    End Select
End Sub

Public Sub ChangeEditableState(ByRef TextBox As TextBox, ByVal Editable As Boolean, Optional ForeColorHighlighted As Boolean = False, Optional ChangeTabStop As Boolean = True)
    Const EDITABLE_BACKCOLOR As Long = vbWindowBackground
    Const EDITABLE_FORECOLOR As Long = vbWindowText
    Const NONEDITABLE_BACKCOLOR As Long = vbButtonFace
    Const NONEDITABLE_FORECOLOR As Long = vbHighlight
    
    TextBox.Locked = Not Editable
    If ChangeTabStop Then
        TextBox.TabStop = Editable
    End If
    If Editable Then
        TextBox.BackColor = EDITABLE_BACKCOLOR
        If ForeColorHighlighted Then
            TextBox.ForeColor = EDITABLE_FORECOLOR
        End If
    Else
        TextBox.BackColor = NONEDITABLE_BACKCOLOR
        If ForeColorHighlighted Then
            TextBox.ForeColor = NONEDITABLE_FORECOLOR
        End If
    End If
End Sub

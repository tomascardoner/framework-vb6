Attribute VB_Name = "CSM_Function"
Option Explicit

'================================================================================================================
' IF IS NULL

Public Function IfIsNull(ByVal CheckValue As Variant, ByVal ReplacementValue As Variant) As Variant
    IfIsNull = IIf(IsNull(CheckValue), ReplacementValue, CheckValue)
End Function

Public Function IfIsNull_Zero(ByVal CheckValue As Variant) As Variant
    IfIsNull_Zero = IIf(IsNull(CheckValue), 0, CheckValue)
End Function

Public Function IfIsNull_Two(ByVal CheckValue As Variant) As Variant
    IfIsNull_Two = IIf(IsNull(CheckValue), 2, CheckValue)
End Function

Public Function IfIsNull_Space(ByVal CheckValue As Variant) As String
    IfIsNull_Space = IIf(IsNull(CheckValue), " ", CheckValue)
End Function

Public Function IfIsNull_ZeroLenghtString(ByVal CheckValue As Variant) As Variant
    IfIsNull_ZeroLenghtString = IIf(IsNull(CheckValue), "", CheckValue)
End Function

Public Function IfIsNull_ZeroDate(ByVal CheckValue As Variant) As Variant
    IfIsNull_ZeroDate = IIf(IsNull(CheckValue), CSM_Constant.DATE_TIME_FIELD_NULL_VALUE, CheckValue)
End Function

'================================================================================================================
' IF IS ZERO

Public Function IfIsZero(ByVal CheckValue As Variant, ByVal ReplacementValue As Variant) As Variant
    IfIsZero = IIf(CheckValue = 0, ReplacementValue, CheckValue)
End Function

Public Function IfIsZero_Null(ByVal CheckValue As Variant) As Variant
    IfIsZero_Null = IIf(CheckValue = 0, Null, CheckValue)
End Function

Public Function IfIsTwo_Null(ByVal CheckValue As Variant) As Variant
    IfIsTwo_Null = IIf(CheckValue = 2, Null, CheckValue)
End Function

Public Function IfIsZeroLenghtString_Null(ByVal CheckValue As String, Optional TrimValue As Boolean = True) As Variant
    If TrimValue Then
        IfIsZeroLenghtString_Null = IIf(Trim(CheckValue) = "", Null, Trim(CheckValue))
    Else
        IfIsZeroLenghtString_Null = IIf(CheckValue = "", Null, CheckValue)
    End If
End Function

Public Function IfIsZeroDate_Null(ByVal CheckValue As Variant) As Variant
    IfIsZeroDate_Null = IIf(CheckValue = CSM_Constant.DATE_TIME_FIELD_NULL_VALUE, Null, CheckValue)
End Function

'================================================================================================================

Public Function ComboboxListIndex2SQLBit(ByVal ListIndex As Integer) As Variant
    Select Case ListIndex
        Case 0
            ComboboxListIndex2SQLBit = Null
        Case 1
            ComboboxListIndex2SQLBit = True
        Case 2
            ComboboxListIndex2SQLBit = False
    End Select
End Function

Public Function SQLBit2ComboboxListIndex(ByVal DBValue As Variant) As Integer
    Select Case DBValue
        Case IsNull(DBValue)
            SQLBit2ComboboxListIndex = 0
        Case True
            SQLBit2ComboboxListIndex = 1
        Case False
            SQLBit2ComboboxListIndex = 2
    End Select
End Function

Public Function SQLBit2CheckBoxValue(ByVal DBValue As Variant) As CheckBoxConstants
    If IsNull(DBValue) Then
        SQLBit2CheckBoxValue = CheckBoxConstants.vbGrayed
    ElseIf CBool(DBValue) Then
        SQLBit2CheckBoxValue = CheckBoxConstants.vbChecked
    Else
        SQLBit2CheckBoxValue = CheckBoxConstants.vbUnchecked
    End If
End Function

Public Function CheckBoxValue2SQLBit(ByVal ByteValue As CheckBoxConstants) As Variant
    Select Case ByteValue
        Case CheckBoxConstants.vbUnchecked
            CheckBoxValue2SQLBit = False
        Case CheckBoxConstants.vbChecked
            CheckBoxValue2SQLBit = True
        Case CheckBoxConstants.vbGrayed
            CheckBoxValue2SQLBit = Null
    End Select
End Function

Public Function ComboboxListIndex2CheckBoxValue(ByVal ListIndex As Integer) As CheckBoxConstants
    Select Case ListIndex
        Case 0
            ComboboxListIndex2CheckBoxValue = CheckBoxConstants.vbGrayed
        Case 1
            ComboboxListIndex2CheckBoxValue = CheckBoxConstants.vbChecked
        Case 2
            ComboboxListIndex2CheckBoxValue = CheckBoxConstants.vbUnchecked
    End Select
End Function

Public Function CheckBoxValue2ComboboxListIndex(ByVal CheckBoxValue As CheckBoxConstants) As Integer
    Select Case CheckBoxValue
        Case CheckBoxConstants.vbGrayed
            CheckBoxValue2ComboboxListIndex = 0
        Case CheckBoxConstants.vbChecked
            CheckBoxValue2ComboboxListIndex = 1
        Case CheckBoxConstants.vbUnchecked
            CheckBoxValue2ComboboxListIndex = 2
    End Select
End Function

Public Function BooleanToChecked(ByVal value As Boolean) As CheckBoxConstants
    BooleanToChecked = IIf(value, vbChecked, vbUnchecked)
End Function

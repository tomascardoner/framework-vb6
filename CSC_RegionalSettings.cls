VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_RegionalSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

'Number Format Constants
Private Const LOCALE_SDECIMAL = &HE
Private Const LOCALE_IDIGITS = &H11
Private Const LOCALE_STHOUSAND = &HF
Private Const LOCALE_SGROUPING = &H10
Private Const LOCALE_SNEGATIVESIGN = &H51
Private Const LOCALE_SLIST = &HC

'Currency Format Constants
Private Const LOCALE_SCURRENCY = &H14
Private Const LOCALE_SMONDECIMALSEP = &H16
Private Const LOCALE_ICURRDIGITS = &H19
Private Const LOCALE_SMONTHOUSANDSEP = &H17
Private Const LOCALE_SMONGROUPING = &H18
'=========================================================================

'Number Format Variables
Private mNumberDecimalSymbol As String
Private mNumberNumberOfDigitsAfterDecimal As Byte
Private mNumberDigitGroupingSymbol As String
Private mNumberNumberOfDigitsInGroup As Byte
Private mNumberNegativeSignSymbol As String
Private mNumberListSeparator As String

'Currency Format Variables
Private mCurrencyCurrencySymbol As String
Private mCurrencyDecimalSymbol As String
Private mCurrencyNumberOfDigitsAfterDecimal As Byte
Private mCurrencyDigitGroupingSymbol As String
Private mCurrencyNumberOfDigitsInGroup As Byte
'=========================================================================

Private Function GetValue(ByVal Constant As Long) As String
      Dim Value As String
      Dim iRet1 As Long
      Dim iRet2 As Long
      Dim lpLCDataVar As String
      Dim Pos As Integer
      Dim Locale As Long

      Locale = GetUserDefaultLCID()
      iRet1 = GetLocaleInfo(Locale, Constant, lpLCDataVar, 0)

      Value = String(iRet1, 0)
      iRet2 = GetLocaleInfo(Locale, Constant, Value, iRet1)
      Pos = InStr(Value, Chr(0))
      If Pos > 0 Then
         Value = Left(Value, Pos - 1)
      End If
      GetValue = Value
End Function

'=========================================================================
' Number Format Properties
Public Property Get NumberDecimalSymbol(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mNumberDecimalSymbol = GetValue(LOCALE_SDECIMAL)
    End If
    NumberDecimalSymbol = mNumberDecimalSymbol
End Property

Public Property Get NumberNumberOfDigitsAfterDecimal(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mNumberNumberOfDigitsAfterDecimal = Val(GetValue(LOCALE_IDIGITS))
    End If
    NumberNumberOfDigitsAfterDecimal = mNumberNumberOfDigitsAfterDecimal
End Property

Public Property Get NumberDigitGroupingSymbol(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mNumberDigitGroupingSymbol = GetValue(LOCALE_STHOUSAND)
    End If
    NumberDigitGroupingSymbol = mNumberDigitGroupingSymbol
End Property

Public Property Get NumberNumberOfDigitsInGroup(Optional Refresh As Boolean = False) As Byte
    If Refresh Then
        mNumberNumberOfDigitsInGroup = Val(GetValue(LOCALE_SGROUPING))
    End If
    NumberNumberOfDigitsInGroup = mNumberNumberOfDigitsInGroup
End Property

Public Property Get NumberNegativeSignSymbol(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mNumberNegativeSignSymbol = GetValue(LOCALE_SNEGATIVESIGN)
    End If
    NumberNegativeSignSymbol = mNumberNegativeSignSymbol
End Property

Public Property Get NumberListSeparator(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mNumberListSeparator = GetValue(LOCALE_SLIST)
    End If
    NumberListSeparator = mNumberListSeparator
End Property

'=========================================================================
'Currency Format Properties
Public Property Get CurrencyCurrencySymbol(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mCurrencyCurrencySymbol = GetValue(LOCALE_SCURRENCY)
    End If
    CurrencyCurrencySymbol = mCurrencyCurrencySymbol
End Property

Public Property Get CurrencyNumberOfDigitsAfterDecimal(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mCurrencyNumberOfDigitsAfterDecimal = Val(GetValue(LOCALE_ICURRDIGITS))
    End If
    CurrencyNumberOfDigitsAfterDecimal = mCurrencyNumberOfDigitsAfterDecimal
End Property

Public Property Get CurrencyDecimalSymbol(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mCurrencyDecimalSymbol = GetValue(LOCALE_SMONDECIMALSEP)
    End If
    CurrencyDecimalSymbol = mCurrencyDecimalSymbol
End Property

Public Property Get CurrencyDigitGroupingSymbol(Optional Refresh As Boolean = False) As String
    If Refresh Then
        mCurrencyDigitGroupingSymbol = GetValue(LOCALE_SMONTHOUSANDSEP)
    End If
    CurrencyDigitGroupingSymbol = mCurrencyDigitGroupingSymbol
End Property

Public Property Get CurrencyNumberOfDigitsInGroup(Optional Refresh As Boolean = False) As Byte
    If Refresh Then
        mCurrencyNumberOfDigitsInGroup = Val(GetValue(LOCALE_SMONGROUPING))
    End If
    CurrencyNumberOfDigitsInGroup = mCurrencyNumberOfDigitsInGroup
End Property

'=========================================================================
Private Sub Class_Initialize()
    Dim DummyString As String
    Dim DummyNumber As Double
    
    'Number Format
    DummyString = NumberDecimalSymbol(True)
    DummyNumber = NumberNumberOfDigitsAfterDecimal(True)
    DummyString = NumberDigitGroupingSymbol(True)
    DummyNumber = NumberNumberOfDigitsInGroup(True)
    DummyString = NumberNegativeSignSymbol(True)
    DummyString = NumberListSeparator(True)
    
    'Currency Format
    DummyString = CurrencyCurrencySymbol(True)
    DummyString = CurrencyDecimalSymbol(True)
    DummyNumber = CurrencyNumberOfDigitsAfterDecimal(True)
    DummyString = CurrencyDigitGroupingSymbol(True)
    DummyNumber = CurrencyNumberOfDigitsInGroup(True)
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_ReportParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mParameterName As String
Private mParameterDescription As String
Private mParameterDataType As csReportParameterDataType
Private mParameterAllowMoreThanOne As Boolean
Private mParameterAllowRange As Boolean
Private mParameterRequired As Boolean
Private mParameterListValuesOrRecordSource As String
Private mParameterListFieldNameBound As String
Private mParameterListFieldNameDisplay As String
Private mParameterListErrorEntityName As String
Private mParameterMinValue As Variant
Private mParameterMaxValue As Variant
Private mParameterValue As Variant
Private mParameterDisplayValue As String
Private mParameterAskFor As Boolean

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterName() As String
    ParameterName = mParameterName
End Property

Public Property Let ParameterName(ByVal Value As String)
    mParameterName = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterDescription() As String
    ParameterDescription = mParameterDescription
End Property

Public Property Let ParameterDescription(ByVal Value As String)
    mParameterDescription = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterDataType() As csReportParameterDataType
    ParameterDataType = mParameterDataType
End Property

Public Property Let ParameterDataType(ByVal Value As csReportParameterDataType)
    mParameterDataType = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterAllowMoreThanOne() As Boolean
    ParameterAllowMoreThanOne = mParameterAllowMoreThanOne
End Property

Public Property Let ParameterAllowMoreThanOne(ByVal Value As Boolean)
    mParameterAllowMoreThanOne = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterAllowRange() As Boolean
    ParameterAllowRange = mParameterAllowRange
End Property

Public Property Let ParameterAllowRange(ByVal Value As Boolean)
    mParameterAllowRange = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterRequired() As Boolean
    ParameterRequired = mParameterRequired
End Property

Public Property Let ParameterRequired(ByVal Value As Boolean)
    mParameterRequired = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterListValuesOrRecordSource() As String
    ParameterListValuesOrRecordSource = mParameterListValuesOrRecordSource
End Property

Public Property Let ParameterListValuesOrRecordSource(ByVal Value As String)
    mParameterListValuesOrRecordSource = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterListFieldNameBound() As String
    ParameterListFieldNameBound = mParameterListFieldNameBound
End Property

Public Property Let ParameterListFieldNameBound(ByVal Value As String)
    mParameterListFieldNameBound = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterListFieldNameDisplay() As String
    ParameterListFieldNameDisplay = mParameterListFieldNameDisplay
End Property

Public Property Let ParameterListFieldNameDisplay(ByVal Value As String)
    mParameterListFieldNameDisplay = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterListErrorEntityName() As String
    ParameterListErrorEntityName = mParameterListErrorEntityName
End Property

Public Property Let ParameterListErrorEntityName(ByVal Value As String)
    mParameterListErrorEntityName = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterMinValue() As Variant
    ParameterMinValue = mParameterMinValue
End Property

Public Property Let ParameterMinValue(ByVal Value As Variant)
    mParameterMinValue = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterMaxValue() As Variant
    ParameterMaxValue = mParameterMaxValue
End Property

Public Property Let ParameterMaxValue(ByVal Value As Variant)
    mParameterMaxValue = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterValue() As Variant
    ParameterValue = mParameterValue
End Property

Public Property Let ParameterValue(ByVal Value As Variant)
    mParameterValue = Value
    Select Case mParameterDataType
        Case csrpdtUndefined
        Case csrpdtString
        Case csrpdtNumberInteger
        Case csrpdtNumberDecimal
        Case csrpdtCurrency
        Case csrpdtDate
        Case csrpdtTime
        Case csrpdtDateTime
        Case csrpdtBoolean
        Case csrpdtList
        Case csrpdtWeekday
            If mParameterDisplayValue = "" Then
                mParameterDisplayValue = WeekdayName(mParameterValue)
            End If
        Case csrpdtMonth
            If mParameterDisplayValue = "" And Not IsEmpty(mParameterValue) Then
                mParameterDisplayValue = MonthName(mParameterValue)
            End If
        Case csrpdtYear
    End Select
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterDisplayValue() As String
    ParameterDisplayValue = mParameterDisplayValue
End Property

Public Property Let ParameterDisplayValue(ByVal Value As String)
    mParameterDisplayValue = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParameterAskFor() As Boolean
    ParameterAskFor = mParameterAskFor
End Property

Public Property Let ParameterAskFor(ByVal Value As Boolean)
    mParameterAskFor = Value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub Class_Initialize()
    mParameterName = ""
    mParameterDescription = ""
    mParameterDataType = csrpdtUndefined
    mParameterAllowMoreThanOne = False
    mParameterAllowRange = False
    mParameterRequired = False
    mParameterListValuesOrRecordSource = ""
    mParameterListFieldNameBound = ""
    mParameterListFieldNameDisplay = ""
    mParameterListErrorEntityName = ""
    mParameterMinValue = Empty
    mParameterMaxValue = Empty
    mParameterValue = Empty
    mParameterDisplayValue = ""
    mParameterAskFor = True
End Sub

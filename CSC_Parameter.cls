VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_Parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDatabaseTableKeepOpen As Boolean

Private mTableName As String

Private mTableFieldNamePrimaryKey As String
Private mTableFieldNameText As String
Private mTableFieldNameNumberInteger As String
Private mTableFieldNameNumberDecimal As String
Private mTableFieldNameCurrency As String
Private mTableFieldNameDateTime As String
Private mTableFieldNameBoolean As String

Private mrecParametro As ADODB.Recordset

Public Enum CSParameterDataTypes
    cspdtString
    cspdtNumberInteger
    cspdtNumberDecimal
    cspdtCurrency
    cspdtDateTime
    cspdtBoolean
End Enum

Public Property Get TableName() As String
    TableName = mTableName
End Property

Public Property Let TableName(ByVal Value As String)
    mTableName = Value
End Property

Public Property Get TableFieldNamePrimaryKey() As String
    TableFieldNamePrimaryKey = mTableFieldNamePrimaryKey
End Property

Public Property Let TableFieldNamePrimaryKey(ByVal Value As String)
    mTableFieldNamePrimaryKey = Value
End Property

Public Property Get TableFieldNameText() As String
    TableFieldNameText = mTableFieldNameText
End Property

Public Property Let TableFieldNameText(ByVal Value As String)
    mTableFieldNameText = Value
End Property

Public Property Get TableFieldNameNumberInteger() As String
    TableFieldNameNumberInteger = mTableFieldNameNumberInteger
End Property

Public Property Let TableFieldNameNumberInteger(ByVal Value As String)
    mTableFieldNameNumberInteger = Value
End Property

Public Property Get TableFieldNameNumberDecimal() As String
    TableFieldNameNumberDecimal = mTableFieldNameNumberDecimal
End Property

Public Property Let TableFieldNameNumberDecimal(ByVal Value As String)
    mTableFieldNameNumberDecimal = Value
End Property

Public Property Get TableFieldNameCurrency() As String
    TableFieldNameCurrency = mTableFieldNameCurrency
End Property

Public Property Let TableFieldNameCurrency(ByVal Value As String)
    mTableFieldNameCurrency = Value
End Property

Public Property Get TableFieldNameDateTime() As String
    TableFieldNameDateTime = mTableFieldNameDateTime
End Property

Public Property Let TableFieldNameDateTime(ByVal Value As String)
    mTableFieldNameDateTime = Value
End Property

Public Property Get TableFieldNameBoolean() As String
    TableFieldNameBoolean = mTableFieldNameBoolean
End Property

Public Property Let TableFieldNameBoolean(ByVal Value As String)
    mTableFieldNameBoolean = Value
End Property


Private Function OpenDatabaseParameter() As Boolean
    If mrecParametro Is Nothing Then
        OpenDatabaseParameter = pDatabase.OpenRecordset(mrecParametro, "SELECT * FROM " & mTableName, adOpenStatic, adLockOptimistic, adCmdText, "Error al abrir la Tabla de Parámetros.", "Classes.CSC_Parameter.OpenDatabaseParameter")
    ElseIf mrecParametro.State = adStateClosed Then
        OpenDatabaseParameter = pDatabase.OpenRecordset(mrecParametro, "SELECT * FROM " & mTableName, adOpenStatic, adLockOptimistic, adCmdText, "Error al abrir la Tabla de Parámetros.", "Classes.CSC_Parameter.OpenDatabaseParameter")
    Else
        OpenDatabaseParameter = True
    End If
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetParameter(ByVal ParameterID As String, ByVal ParameterDataType As CSParameterDataTypes, Optional ByVal DefaultValue As Variant) As Variant
    Dim FieldName As String

    Select Case ParameterDataType
        Case cspdtString
            FieldName = mTableFieldNameText
        Case cspdtNumberInteger
            FieldName = mTableFieldNameNumberInteger
        Case cspdtNumberDecimal
            FieldName = mTableFieldNameNumberDecimal
        Case cspdtCurrency
            FieldName = mTableFieldNameCurrency
        Case cspdtDateTime
            FieldName = mTableFieldNameDateTime
        Case cspdtBoolean
            FieldName = mTableFieldNameBoolean
    End Select
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If OpenDatabaseParameter() Then
        mrecParametro.Filter = mTableFieldNamePrimaryKey & " = '" & ParameterID & "'"
        If mrecParametro.EOF Then
            GetParameter = DefaultValue
        ElseIf IsNull(mrecParametro(FieldName).Value) Then
            GetParameter = DefaultValue
        Else
            GetParameter = mrecParametro(FieldName).Value
        End If
        
        If Not mDatabaseTableKeepOpen Then
            mrecParametro.Close
            Set mrecParametro = Nothing
        End If
    End If
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.CSC_Parameter.GetParameter", "Error al obtener el Parámetro." & vbCr & vbCr & "IDParametro: " & ParameterID
End Function

Public Function SetParameter(ByVal ParameterID As String, ByVal ParameterDataType As CSParameterDataTypes, ByVal Value As Variant) As Boolean
    Dim FieldName As String

    Select Case ParameterDataType
        Case cspdtString
            FieldName = mTableFieldNameText
        Case cspdtNumberInteger
            FieldName = mTableFieldNameNumberInteger
        Case cspdtNumberDecimal
            FieldName = mTableFieldNameNumberDecimal
        Case cspdtCurrency
            FieldName = mTableFieldNameCurrency
        Case cspdtDateTime
            FieldName = mTableFieldNameDateTime
        Case cspdtBoolean
            FieldName = mTableFieldNameBoolean
    End Select
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If OpenDatabaseParameter() Then
        mrecParametro.Filter = mTableFieldNamePrimaryKey & " = '" & ParameterID & "'"
        If mrecParametro.EOF Then
            mrecParametro.AddNew
            mrecParametro(mTableFieldNamePrimaryKey).Value = ParameterID
        End If
        mrecParametro(FieldName).Value = Value
        mrecParametro.Update
        
        If Not mDatabaseTableKeepOpen Then
            mrecParametro.Close
            Set mrecParametro = Nothing
        End If
        
        SetParameter = True
    End If
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CSC_Parameter.SetParameter", "Error al guardar el Parámetro." & vbCr & vbCr & "IDParametro: " & ParameterID
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetParameterText(ByVal ParameterID As String, Optional ByVal DefaultValue As String = "") As String
    GetParameterText = GetParameter(ParameterID, cspdtString, DefaultValue)
End Function

Public Function SetParameterText(ByVal ParameterID As String, ByVal Value As String) As Boolean
    SetParameterText = SetParameter(ParameterID, cspdtString, Value)
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetParameterCollection(ByVal ParameterID As String, Optional ByVal DefaultValue As String = "") As Collection
    Dim ValueList As String
    Dim aValues() As String

    ValueList = GetParameterText(ParameterID, DefaultValue)
    If Trim(ValueList) <> "" Then
        aValues = Split(ValueList, ";")
        Set GetParameterCollection = CSM_String.ConvertArrayToCollection(aValues)
    End If
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetParameterNumberInteger(ByVal ParameterID As String, Optional ByVal DefaultValue As Long = 0) As Long
    GetParameterNumberInteger = GetParameter(ParameterID, cspdtNumberInteger, DefaultValue)
End Function

Public Function SetParameterNumberInteger(ByVal ParameterID As String, ByVal Value As Long) As Boolean
    SetParameterNumberInteger = SetParameter(ParameterID, cspdtNumberInteger, Value)
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetParameterNumberDecimal(ByVal ParameterID As String, Optional ByVal DefaultValue As Double = 0) As Double
    GetParameterNumberDecimal = GetParameter(ParameterID, cspdtNumberDecimal, DefaultValue)
End Function

Public Function SetParameterNumberDecimal(ByVal ParameterID As String, ByVal Value As Double) As Boolean
    SetParameterNumberDecimal = SetParameter(ParameterID, cspdtNumberDecimal, Value)
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetParameterCurrency(ByVal ParameterID As String, Optional ByVal DefaultValue As Currency = 0) As Currency
    GetParameterCurrency = GetParameter(ParameterID, cspdtCurrency, DefaultValue)
End Function

Public Function SetParameterCurrency(ByVal ParameterID As String, ByVal Value As Currency) As Boolean
    SetParameterCurrency = SetParameter(ParameterID, cspdtCurrency, Value)
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetParameterDateTime(ByVal ParameterID As String, Optional ByVal DefaultValue As Date = #12:00:00 AM#) As Date
    GetParameterDateTime = GetParameter(ParameterID, cspdtDateTime, DefaultValue)
End Function

Public Function SetParameterDateTime(ByVal ParameterID As String, ByVal Value As Date) As Boolean
    SetParameterDateTime = SetParameter(ParameterID, cspdtDateTime, Value)
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetParameterBoolean(ByVal ParameterID As String, Optional ByVal DefaultValue As Boolean = False) As Boolean
    GetParameterBoolean = GetParameter(ParameterID, cspdtBoolean, DefaultValue)
End Function

Public Function SetParameterBoolean(ByVal ParameterID As String, ByVal Value As Boolean) As Boolean
    SetParameterBoolean = SetParameter(ParameterID, cspdtBoolean, Value)
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub Class_Initialize()
    mTableName = "Parametro"
    
    mTableFieldNamePrimaryKey = "IDParametro"
    mTableFieldNameText = "Texto"
    mTableFieldNameNumberInteger = "NumeroEntero"
    mTableFieldNameNumberDecimal = "NumeroDecimal"
    mTableFieldNameCurrency = "Moneda"
    mTableFieldNameDateTime = "FechaHora"
    mTableFieldNameBoolean = "SiNo"
End Sub

Private Sub Class_Terminate()
    If Not mrecParametro Is Nothing Then
        If mrecParametro.State = adStateOpen Then
            mrecParametro.Close
        End If
        Set mrecParametro = Nothing
    End If
End Sub

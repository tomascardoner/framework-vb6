VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_DatabaseModel_ADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMasterConnection As ADODB.Connection
Private mDatabases As Collection

Public Property Get Databases() As Collection
    If mDatabases Is Nothing Then
        Call LoadDatabases
    End If
    Set Databases = mDatabases
End Property

Public Function Connect() As Boolean
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set mMasterConnection = New ADODB.Connection
    With mMasterConnection
        .Provider = pParametro.Database_Provider
        .ConnectionTimeout = pParametro.Database_ConnectionTimeout
        .CommandTimeout = pParametro.Database_CommandTimeout
        .CursorLocation = pParametro.Database_CursorLocation
        .Mode = adModeUnknown
        .ConnectionString = "Data Source=" & pParametro.Database_DataSource & ";User ID=" & pParametro.Database_UserID & ";Password=" & pParametro.Database_Password & ";Packet Size=" & pParametro.Database_PacketSize & ";DataTypeCompatibility=80;MARS Connection=True"
        .Open
        .DefaultDatabase = "master"
    End With
        
    Connect = True
    
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        If pDatabase.Connection.Errors(0).NativeError = 945 Then
            ShowErrorMessage "Classes.CSC_DatabaseModel_ADO.Connect", "No existe la Base de Datos." & vbCr & vbCr & "Data Source: " & pParametro.Database_DataSource & vbCr & "Database: " & pParametro.Database_Database
        Else
            ShowErrorMessage "Classes.CSC_DatabaseModel_ADO.Connect", "Error al crear la conexión al Servidor de Base de Datos." & vbCr & vbCr & "Data Source: " & pParametro.Database_DataSource
        End If
    Else
        ShowErrorMessage "Classes.CSC_DatabaseModel_ADO.Connect", "Error al crear la conexión al Servidor de Base de Datos." & vbCr & vbCr & "Data Source: " & pParametro.Database_DataSource
    End If
End Function

Public Function LoadDatabases() As Boolean
    Dim recCatalogs As ADODB.Recordset
    
    If mMasterConnection Is Nothing Then
        If Not Connect() Then
            Exit Function
        End If
    End If
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Set recCatalogs = New ADODB.Recordset
    Set recCatalogs.ActiveConnection = mMasterConnection
    recCatalogs.CursorType = adOpenForwardOnly
    recCatalogs.LockType = adLockReadOnly
    recCatalogs.Source = "SELECT "
    recCatalogs.Open
    
    Set mDatabases = pDatabase.Connection.ca

ErrorHandler:
    If pDatabase.Connection.Errors.Count > 0 Then
        ShowErrorMessage "Classes.CSC_DatabaseModel_ADO.LoadDatabases", "Error al cargar la lista de Bases de Datos."
    End If
End Function

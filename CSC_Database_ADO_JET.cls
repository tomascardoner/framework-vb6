VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_Database_ADO_JET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'///////////////////////////////////////////////////////////////////
'DATABASE CONNECTION ADO
Private mProvider As String
Private mConnectionTimeout As Long
Private mCommandTimeout As Long
Private mCDataSourcesNames As Collection
Private mCDataSources As Collection
Private mDataSource As String
Private mUserID As String
Private mPassword  As String
Private mDatabase As String
Private mDatabasePassword As String
Private mCursorLocationServer As Boolean
Private mOpenExclusive As Boolean
Private mReadOnly As Boolean

Private mReportsPath As String

Private mBackupCopiesNumber As Long

Private mCompactDays As Long
Private mCompactLastDate As Date

Private mParametersSaveToRegistry As Boolean

Private mConnection As ADODB.Connection

'///////////////////////////////////////////////////////
'ERRORS MEMBERS

Public Property Get ERROR_INVALID_PASSWORD() As Long
    ERROR_INVALID_PASSWORD = -124782449
End Property

Public Property Get ERROR_DUPLICATE_KEY() As Long
    ERROR_DUPLICATE_KEY = -105121349
End Property

Public Property Get ERROR_RELATED_RECORDS() As Long
    ERROR_RELATED_RECORDS = -534971980
End Property

'///////////////////////////////////////////////////////

Public Property Get Provider() As String
    Provider = mProvider
End Property

Public Property Let Provider(ByVal Value As String)
    mProvider = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get ConnectionTimeout() As Long
    ConnectionTimeout = mConnectionTimeout
End Property

Public Property Let ConnectionTimeout(ByVal Value As Long)
    mConnectionTimeout = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get CommandTimeout() As Long
    CommandTimeout = mCommandTimeout
End Property

Public Property Let CommandTimeout(ByVal Value As Long)
    mCommandTimeout = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get CDataSourcesNames() As Collection
    Set CDataSourcesNames = mCDataSourcesNames
End Property

Public Property Let CDataSourcesNames(ByVal Value As Collection)
    Set mCDataSourcesNames = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get CDataSources() As Collection
    Set CDataSources = mCDataSources
End Property

Public Property Let CDataSources(ByVal Value As Collection)
    Set mCDataSources = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get DataSource() As String
    DataSource = mDataSource
End Property

Public Property Let DataSource(ByVal Value As String)
    mDataSource = Value
End Property

Public Property Get DataSourceFolder() As String
    DataSourceFolder = CSM_String.RemoveLastSubString(DataSource, "\")
End Property

'///////////////////////////////////////////////////////

Public Property Get UserID() As String
    UserID = mUserID
End Property

Public Property Let UserID(ByVal Value As String)
    mUserID = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(ByVal Value As String)
    mPassword = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get Database() As String
    Database = mDatabase
End Property

Public Property Let Database(ByVal Value As String)
    mDatabase = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get DatabasePassword() As String
    DatabasePassword = mDatabasePassword
End Property

Public Property Let DatabasePassword(ByVal Value As String)
    mDatabasePassword = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get CursorLocationServer() As Boolean
    CursorLocationServer = mCursorLocationServer
End Property

Public Property Let CursorLocationServer(ByVal Value As Boolean)
    mCursorLocationServer = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get OpenExclusive() As Boolean
    OpenExclusive = mOpenExclusive
End Property

Public Property Let OpenExclusive(ByVal Value As Boolean)
    mOpenExclusive = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property

Public Property Let ReadOnly(ByVal Value As Boolean)
    mReadOnly = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get ReportsPath() As String
    ReportsPath = mReportsPath
End Property

Public Property Let ReportsPath(ByVal Value As String)
    mReportsPath = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get BackupCopiesNumber() As Long
    BackupCopiesNumber = mBackupCopiesNumber
End Property

Public Property Let BackupCopiesNumber(ByVal Value As Long)
    mBackupCopiesNumber = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get ParametersSaveToRegistry() As Boolean
    ParametersSaveToRegistry = mParametersSaveToRegistry
End Property

Public Property Let ParametersSaveToRegistry(ByVal Value As Boolean)
    mParametersSaveToRegistry = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get CompactDays() As Long
    CompactDays = mCompactDays
End Property

Public Property Let CompactDays(ByVal Value As Long)
    mCompactDays = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get CompactLastDate() As Date
    CompactLastDate = mCompactLastDate
End Property

Public Property Let CompactLastDate(ByVal Value As Date)
    mCompactLastDate = Value
End Property

'///////////////////////////////////////////////////////

Public Function LoadParameters() As Boolean
    If mParametersSaveToRegistry Then
        LoadParameters = LoadParametersFromRegistry()
    Else
        LoadParameters = LoadParametersFromINI()
    End If
End Function

Public Function LoadParametersFromRegistry() As Boolean
    Dim DES As CSC_Encryption_DES
    Dim CValuesNames As Collection
    Dim CValuesTypes As Collection
    Dim CValues As Collection
    Dim index As Integer
    
    mProvider = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "Provider", "", csrdtString)
    mConnectionTimeout = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "ConnectionTimeout", 30, csrdtNumberInteger)
    mCommandTimeout = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "CommandTimeout", 30, csrdtNumberInteger)
    
    'NEW FEATURE: MULTIPLE DATASOURCES
    'mDataSource = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "DataSource", "", csrdtString)
    Call CSM_Registry.EnumerateValues_FromApplication_LocalMachine("Database", CValuesNames, CValuesTypes, CValues, "DataSource", False)
    Select Case CValuesNames.Count
        Case 0 'No Datasource specified
            Set mCDataSourcesNames = New Collection
            Set mCDataSources = New Collection
            mDataSource = ""
        Case 1
            mCDataSourcesNames.Add CValuesNames(1)
            mCDataSources.Add CValues(1)
            mDataSource = CValues(1)
        Case Else
            For index = 1 To CValuesNames.Count
                mCDataSourcesNames.Add Trim(Replace(Mid(CValuesNames(index), Len("DataSource") + 1), "_", " "))
                mCDataSources.Add CValues(index)
            Next index
    End Select
    Set CValuesNames = Nothing
    Set CValuesTypes = Nothing
    Set CValues = Nothing
    
    mUserID = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "UserID", "", csrdtString)
    mPassword = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "Password", "", csrdtString)
    If mPassword <> "" Then
        Set DES = New CSC_Encryption_DES
        mPassword = DES.DecryptString(mPassword, DES.PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    mDatabase = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "Database", "", csrdtString)
    mDatabasePassword = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "DatabasePassword", "", csrdtString)
    If mDatabasePassword <> "" Then
        Set DES = New CSC_Encryption_DES
        mDatabasePassword = DES.DecryptString(mDatabasePassword, DES.PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    
    mReportsPath = CSM_Registry.GetValue_FromApplication_LocalMachine("Reports", "Path", "", csrdtString)
    
    mBackupCopiesNumber = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "BackupCopiesNumber", 7, csrdtNumberInteger)
    
    LoadParametersFromRegistry = True
End Function

Public Function LoadParametersFromINI() As Boolean
    Dim DES As CSC_Encryption_DES
    
    mProvider = CSM_INI.GetValue_FromApplication("Database", "Provider", "", csidtString)
    mConnectionTimeout = CSM_INI.GetValue_FromApplication("Database", "ConnectionTimeout", 15, csidtNumberInteger)
    mCommandTimeout = CSM_INI.GetValue_FromApplication("Database", "CommandTimeout", 15, csidtNumberInteger)
    mDataSource = CSM_INI.GetValue_FromApplication("Database", "DataSource", "", csidtString)
    mUserID = CSM_INI.GetValue_FromApplication("Database", "UserID", "", csidtString)
    mPassword = CSM_INI.GetValue_FromApplication("Database", "Password", "", csidtString)
    If mPassword <> "" Then
        Set DES = New CSC_Encryption_DES
        mPassword = DES.DecryptString(mPassword, DES.PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    mDatabase = CSM_INI.GetValue_FromApplication("Database", "Database", "", csidtString)
    mDatabasePassword = CSM_INI.GetValue_FromApplication("Database", "DatabasePassword", "", csidtString)
    If mDatabasePassword <> "" Then
        Set DES = New CSC_Encryption_DES
        mDatabasePassword = DES.DecryptString(mDatabasePassword, DES.PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    
    mReportsPath = CSM_INI.GetValue_FromApplication("Reports", "Path", "", csidtString)
    
    mBackupCopiesNumber = CSM_INI.GetValue_FromApplication("Database", "BackupCopiesNumber", 7, csidtNumberInteger)

    LoadParametersFromINI = True
End Function

'///////////////////////////////////////////////////////

Public Function SaveParameters() As Boolean
    If mParametersSaveToRegistry Then
        SaveParameters = SaveParametersToRegistry
    Else
        SaveParameters = SaveParametersToINI
    End If
End Function

Public Function SaveParametersToRegistry() As Boolean
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "Provider", mProvider)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "ConnectionTimeout", mConnectionTimeout)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "CommandTimeout", mCommandTimeout)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "DataSource", mDataSource)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "UserID", mUserID)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "Password", mPassword)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "Name", mDatabase)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "BackupCopiesNumber", mBackupCopiesNumber)
    
    SaveParametersToRegistry = True
End Function

Public Function SaveParametersToINI() As Boolean
    Call CSM_INI.SetValue_ToApplication("Database", "Provider", mProvider)
    Call CSM_INI.SetValue_ToApplication("Database", "ConnectionTimeout", mConnectionTimeout)
    Call CSM_INI.SetValue_ToApplication("Database", "CommandTimeout", mCommandTimeout)
    Call CSM_INI.SetValue_ToApplication("Database", "DataSource", mDataSource)
    Call CSM_INI.SetValue_ToApplication("Database", "UserID", mUserID)
    Call CSM_INI.SetValue_ToApplication("Database", "Password", mPassword)
    Call CSM_INI.SetValue_ToApplication("Database", "Name", mDatabase)
    Call CSM_INI.SetValue_ToApplication("Database", "BackupCopiesNumber", mBackupCopiesNumber)
    
    SaveParametersToINI = True
End Function

Public Property Get Connection() As ADODB.Connection
    Set Connection = mConnection
End Property

Public Function Connect() As Long

    Screen.MousePointer = vbHourglass
    
    CSM_ApplicationLog.WriteLogEvent "Database Connection: INIT", vbLogEventTypeInformation, pLogEnabled

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    CSM_ApplicationLog.WriteLogEvent "Database Connection: Create ADODB Object", vbLogEventTypeInformation, pLogEnabled
    Set mConnection = New ADODB.Connection
    
    CSM_ApplicationLog.WriteLogEvent "Database Connection: Setting ADODB Object Properties", vbLogEventTypeInformation, pLogEnabled
    With mConnection
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionTimeout = mConnectionTimeout
        .CommandTimeout = mCommandTimeout
        .CursorLocation = IIf(mCursorLocationServer, adUseServer, adUseClient)
        .Mode = IIf(mOpenExclusive, adModeShareExclusive, adModeShareDenyNone) Or IIf(mReadOnly, adModeRead, adModeReadWrite)
        .Properties("Data Source").Value = mDataSource
        .Properties("Persist Security Info").Value = False
        If mDatabasePassword <> "" Then
            .Properties("Jet OLEDB:Database Password").Value = mDatabasePassword
        End If
        .Properties("User ID").Value = mUserID
        .Properties("Password").Value = mPassword
        .Open
        If mDatabase <> "" Then
            .DefaultDatabase = mDatabase
        End If
    End With
    
    Connect = -1
    CSM_ApplicationLog.WriteLogEvent "Database Connection: DONE", vbLogEventTypeInformation, pLogEnabled
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    If mConnection.Errors.Count > 0 Then
        Connect = mConnection.Errors(0).NativeError
        If mConnection.Errors(0).NativeError = ERROR_INVALID_PASSWORD Then
            Exit Function
        End If
    End If
    CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_JET.Connect", "Error al crear la conexión a la Base de Datos." & vbCr & vbCr & "Data Source: " & mDataSource
End Function

Public Function LoadParametersFromDatabase() As Boolean
    mCompactDays = pCSC_Parameter.GetParameterNumberInteger("DatabaseCompactDays", 7)
    mCompactLastDate = pCSC_Parameter.GetParameterDateTime("DatabaseCompactLastDate", Date)
    
    LoadParametersFromDatabase = True
End Function

Public Function CheckDatabase() As Boolean
    Dim recData As ADODB.Recordset

    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set recData = New ADODB.Recordset
    
    Set recData.ActiveConnection = mConnection
    recData.Source = "SELECT Texto FROM Parametro WHERE IDParametro = 'DatabaseIdentifier'"
    recData.CursorType = adOpenForwardOnly
    recData.LockType = adLockReadOnly
    recData.Open , , , , adCmdText
    
    If Not recData.EOF Then
        If recData("Texto").Value = DATABASE_IDENTIFIER Then
            recData.Close
            Set recData = Nothing
            CheckDatabase = True
            Screen.MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    recData.Close
    Set recData = Nothing
    
ErrorHandler:
    Screen.MousePointer = vbDefault
    MsgBox "La Base de Datos no corresponde a esta Aplicación." & vbCr & vbCr & "Origen de los Datos: " & mDataSource, vbExclamation, App.Title
End Function

Public Function Disconnect() As Boolean
RETRY:
    If Not mConnection Is Nothing Then
        WriteLogEvent "Closing Database Connection", vbLogEventTypeInformation, pLogEnabled
        Select Case mConnection.State
            Case adStateClosed
            Case adStateOpen
                mConnection.Close
            Case adStateConnecting, adStateExecuting, adStateFetching
                DoEvents
                GoTo RETRY
        End Select
        
        Set mConnection = Nothing
    End If
End Function

Public Function VerifyBackup() As Boolean
    Dim FileName As String
    Dim Path As String
    Dim PathAndFileName As String
    Dim DatabaseCompacted As Boolean
    
    If mBackupCopiesNumber = 0 Then
        VerifyBackup = True
        Exit Function
    End If
    
    FileName = App.EXEName & "_" & Format(Date, "yyyymmdd") & ".mdb"
    Path = App.Path & IIf(Right(App.Path, 2) = "\", "", "\") & "Backup"
    PathAndFileName = Path & "\" & FileName
    
    If FileSystem.Dir(Path, vbDirectory) = "" Then
        'LA CARPETA BACKUP NO EXISTE, LA CREO
        FileSystem.MkDir Path
    End If
        
    If Dir(PathAndFileName, vbArchive) = "" Then
        'NO SE HIZO BACKUP, HACER
        
        'CIERRO LA BASE DE DATOS
        Call Disconnect
        
        'COPIA LA BASE DE DATOS
        CSF_Status.lblStatus.Caption = "Realizando Copia de Seguridad..."
        CSF_Status.Show , frmSplash
        DoEvents
        On Error Resume Next
        FileSystem.FileCopy mDataSource, PathAndFileName
        Select Case Err.Number
            Case 0
                'SE HIZO BACKUP CORRECTAMENTE
                'ELIMINO LOS BACKUPS ANTERIORES
                Call DeleteBackupOld
            Case 70
                'LA BASE ESTA ABIERTA, NO SE PUEDE COPIAR,
                'POR ENDE, NO BORRO LOS BACKUPS ANTERIORES
            Case Else
                ShowErrorMessage "Classes.CSC_Database_ADO.VerifyBackup", "Error al realizar la copia de seguridad de la Base de Datos."
        End Select
        On Error GoTo 0
        
        'COMPACTO LA BASE DE DATOS
        DatabaseCompacted = AutoCompact()
        
        Unload CSF_Status
        Set CSF_Status = Nothing
        
        'VUELVO A ABRIR LA BASE DE DATOS
        VerifyBackup = (Connect() = -1)
        
        If VerifyBackup And DatabaseCompacted Then
            'ESTABLEZCO LA FECHA EN LA QUE SE COMPACTÓ LA BASE DE DATOS
            Call pCSC_Parameter.SetParameterDateTime("DatabaseCompactLastDate", Date)
        End If
    Else
        VerifyBackup = True
    End If
End Function

Private Function DeleteBackupOld() As Boolean
    Dim FileName As String
    Dim Path As String
    Dim PathAndFileName As String
    Dim FileNamesCount As Long
    Dim aFileNames() As String
    Dim index As Long
    Dim TestDate As String
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Path = App.Path & IIf(Right(App.Path, 2) = "\", "", "\") & "Backup"
    
    'PRIMERO LLENO EL ARRAY CON LOS NOMBRES DE LOS ARCHIVOS
    FileName = FileSystem.Dir(Path & "\" & App.EXEName & "_????????.mdb", vbNormal)
    Do While FileName <> ""
        TestDate = Mid(FileName, Len(App.EXEName) + 2, 8)
        If IsNumeric(TestDate) Then
            On Error Resume Next
            TestDate = CDate(Mid(TestDate, 1, 4) & "/" & Mid(TestDate, 5, 2) & "/" & Mid(TestDate, 7, 2))
            If pTrapErrors Then
                On Error GoTo ErrorHandler
            Else
                On Error GoTo 0
            End If
            FileNamesCount = FileNamesCount + 1
            ReDim Preserve aFileNames(1 To FileNamesCount) As String
            aFileNames(FileNamesCount) = FileName
        End If
        FileName = FileSystem.Dir
    Loop
    
    If FileNamesCount > mBackupCopiesNumber Then
        'SI HAY MAS DE LO NECESARIO...
        'ORDENO EL ARRAY
        Call CSM_Sort.Sort_ArrayString_QuickSort(aFileNames(), 1, FileNamesCount)
        'ELIMINO LOS MAS VIEJOS
        For index = 1 To FileNamesCount - mBackupCopiesNumber
            PathAndFileName = Path & "\" & aFileNames(index)
            On Error Resume Next
            FileSystem.Kill PathAndFileName
            On Error GoTo 0
        Next index
    End If
    
    Screen.MousePointer = vbDefault
    DeleteBackupOld = True
    Exit Function
    
ErrorHandler:
    Screen.MousePointer = vbDefault
    MsgBox "No se hand podido eliminar las Copias de Seguridad Anteriores.", vbExclamation, App.Title
End Function

Private Function AutoCompact() As Boolean
    Dim NewFileName As String
    Dim JetEngine As JRO.JetEngine
    
    If mCompactDays = -1 Or mCompactDays = 0 Or DateDiff("d", mCompactLastDate, Date) < mCompactDays Then
        AutoCompact = False
        Exit Function
    End If
    
    On Error Resume Next
    
    CSF_Status.lblStatus.Caption = "Compactando la Base de Datos..."
    
    NewFileName = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & "_Compacted.mdb"
    
    'VERIFICO SI EXISTE EL ARCHIVO DE DESTINO, Y LO BORRO
    If FileSystem.Dir(NewFileName, vbNormal) <> "" Then
        FileSystem.Kill NewFileName
        If Err.Number > 0 Then
            Exit Function
        End If
    End If

    'COMPACTO LA BASE DE DATOS
    Set JetEngine = New JRO.JetEngine
    JetEngine.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDataSource & ";Jet OLEDB:Database Password=" & mDatabasePassword, _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NewFileName & ";Jet OLEDB:Database Password=" & mDatabasePassword & ";Jet OLEDB:Engine Type=5"
    If Err.Number > 0 Then
        Exit Function
    End If
    
    'ELIMINO EL ARCHIVO ORIGINAL
    FileSystem.Kill mDataSource
    If Err.Number > 0 Then
        'NO SE PUDO ELIMINAR, ELIMINO EL DE DESTINO
        FileSystem.Kill NewFileName
        If Err.Number > 0 Then
            Exit Function
        End If
    End If
    
    'COPIO EL ARCHIVO NUEVO A LA UBICACION ORIGINAL
    FileSystem.FileCopy NewFileName, mDataSource
    If Err.Number > 0 Then
        ShowErrorMessage "Classes.Database.Compact", "Error al Copiar la Base de Datos Compactada a la ubicación original."
        Exit Function
    End If
    
    'ELIMINO EL ARCHIVO NUEVO
    FileSystem.Kill NewFileName
    
    AutoCompact = True
End Function

Public Function Compact(ByVal DestinationFileName As String) As Boolean
    Dim JetEngine As JRO.JetEngine
    
    On Error Resume Next
    
    'VERIFICO SI EXISTE EL ARCHIVO DE DESTINO, Y LO BORRO
    If FileSystem.Dir(DestinationFileName, vbNormal) <> "" Then
        FileSystem.Kill DestinationFileName
        If Err.Number > 0 Then
            Exit Function
        End If
    End If

    'COMPACTO LA BASE DE DATOS
    Set JetEngine = New JRO.JetEngine
    JetEngine.CompactDatabase _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDataSource & ";Jet OLEDB:Database Password=" & mDatabasePassword, _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DestinationFileName & ";Jet OLEDB:Database Password=" & mDatabasePassword & ";Jet OLEDB:Engine Type=5"
    If Err.Number > 0 Then
        Exit Function
    End If
        
    Compact = True
End Function

Public Function Replication_Sync(ByVal RemoteDBPath As String, ByVal EjectDrive As Boolean) As Boolean
    Dim JetReplica As JRO.Replica
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    If RemoteDBPath <> "" Then
        Screen.MousePointer = vbHourglass
        
        Set JetReplica = New JRO.Replica
        Set JetReplica.ActiveConnection = pDatabase.Connection
        Call JetReplica.SYNCHRONIZE(RemoteDBPath, jrSyncTypeImpExp, jrSyncModeDirect)
        
        Screen.MousePointer = vbDefault
        
        If JetReplica.ConflictTables.BOF And JetReplica.ConflictTables.EOF Then
            MsgBox "Se han sincronizado satisfactoriamente las bases de datos.", vbInformation, App.Title
        Else
            MsgBox "Se han sincronizado las bases de datos, pero se producieron conflictos.", vbExclamation, App.Title
        End If
        Set JetReplica = Nothing
        
        If EjectDrive Then
            If CSM_Drive.SafelyRemove(Left(RemoteDBPath, 1)) Then
                MsgBox "Ya puede desenchufar el dispositivo USB.", vbInformation, App.Title
            Else
                MsgBox "No se pudo expulsar el dispositivo USB. Intente la expulsar a través de Windows.", vbCritical, App.Title
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_JET.Replication_Sync", "Error al sincronizar las bases de datos." & vbCr & vbCr & "Local DB: " & pDatabase.DataSource & "Remote DB: " & RemoteDBPath
End Function

Public Function OpenRecordset(ByRef Recordset As ADODB.Recordset, ByVal SQLStatement As String, ByVal CursorType As CursorTypeEnum, ByVal LockType As LockTypeEnum, ByVal Options As Long, ByVal ErrorMessage As String, ByVal ErrorSource As String) As Boolean
    Dim ADODBRecordset As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    Set ADODBRecordset = New ADODB.Recordset
    ADODBRecordset.Open SQLStatement, mConnection, CursorType, LockType, Options
    Set Recordset = ADODBRecordset
    Set ADODBRecordset = Nothing
    
    OpenRecordset = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage ErrorSource, ErrorMessage
    On Error Resume Next
    ADODBRecordset.Close
    Set ADODBRecordset = Nothing
End Function

Public Function GetFirstValueFromRecordset(ByVal SQLStatement As String, ByVal ErrorMessage As String, ByVal ErrorSource As String) As Variant
    Dim ADODBRecordset As ADODB.Recordset
    
    If OpenRecordset(ADODBRecordset, SQLStatement, adOpenForwardOnly, adLockReadOnly, adCmdText, ErrorMessage, ErrorSource) Then
        If pTrapErrors Then
            On Error GoTo ErrorHandler
        End If
        
        If Not ADODBRecordset.EOF Then
            If ADODBRecordset.Fields.Count > 0 Then
                GetFirstValueFromRecordset = ADODBRecordset.Fields(0).Value
            End If
        End If
    End If
    
    ADODBRecordset.Close
    Set ADODBRecordset = Nothing
    Exit Function

ErrorHandler:
    CSM_Error.ShowErrorMessage ErrorSource, ErrorMessage
    On Error Resume Next
    ADODBRecordset.Close
    Set ADODBRecordset = Nothing
End Function

Public Function TableExists(ByVal TableName As String) As Boolean
    Dim recTable As ADODB.Recordset
    
    On Error Resume Next
    
    Set recTable = New ADODB.Recordset
    With recTable
        Set .ActiveConnection = pDatabase.Connection
        .Source = "SELECT 1 FROM " & TableName
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        Err.Clear
        .Open , , , , adCmdText
        TableExists = (Err.Number = 0)
        .Close
    End With
    Set recTable = Nothing
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub Class_Initialize()
    Set mCDataSourcesNames = New Collection
    Set mCDataSources = New Collection
    mParametersSaveToRegistry = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set mCDataSourcesNames = Nothing
    Set mCDataSources = Nothing
    If mConnection.State <> adStateClosed Then
        mConnection.Close
    End If
End Sub

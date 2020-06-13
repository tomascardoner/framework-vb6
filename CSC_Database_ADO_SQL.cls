VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_Database_ADO_SQL"
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
Private mPacketSize As Long
Private mDataTypeCompatibility As Byte
Private mDataSource As String
Private mFailoverPartner As String
Private mUserID As String
Private mPassword  As String
Private mDatabase As String
Private mCursorLocationServer As Boolean
Private mOpenExclusive As Boolean
Private mReadOnly As Boolean

Private mReportsPath As String

Private mBackup_RemoteFolder As String
Private mBackup_LocalFolder As String
Private mBackup_FileName As String

Private mParametersSaveToRegistry As Boolean

Private mConnection As ADODB.Connection

Private Const BLOB_BLOCK_SIZE As Long = 4096

'///////////////////////////////////////////////////////
'ONLY FOR COMPATIBILITY WITH ADO JET
Public BackupCopiesNumber As Byte

'///////////////////////////////////////////////////////
'ERRORS MEMBERS
Public Property Get ERROR_DATABASE_NOTEXIST() As Long
    ERROR_DATABASE_NOTEXIST = 911
End Property

Public Property Get ERROR_DUPLICATE_PRIMARYKEY() As Long
    ERROR_DUPLICATE_PRIMARYKEY = -2147217873
End Property

Public Property Get ERRORNATIVE_DUPLICATEKEY_PRIMARYKEY() As Long
    ERRORNATIVE_DUPLICATEKEY_PRIMARYKEY = 2627
End Property

Public Property Get ERROR_DUPLICATEKEY_UNIQUEINDEX() As Long
    ERROR_DUPLICATEKEY_UNIQUEINDEX = -2147217900
End Property

Public Property Get ERROR_DUPLICATE_ALTERNATEKEY() As Long
    ERROR_DUPLICATE_ALTERNATEKEY = 2601
End Property

Public Property Get ERROR_COMMAND_DUPLICATE_ALTERNATEKEY() As Long
    ERROR_COMMAND_DUPLICATE_ALTERNATEKEY = 50000
End Property

Public Property Get ERROR_RELATED_RECORDS() As Long
    ERROR_RELATED_RECORDS = 547
End Property

'Public Property Get ERRORNATIVE_RELATED_RECORDS() As Long
'    ERRORNATIVE_RELATED_RECORDS = 50000
'End Property
'
Public Property Get ERROR_VALUE_CHANGED_SINCE_LAST_READ() As Long
    ERROR_VALUE_CHANGED_SINCE_LAST_READ = 32
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

Public Property Get PacketSize() As Long
    PacketSize = mPacketSize
End Property

Public Property Let PacketSize(ByVal Value As Long)
    mPacketSize = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get DataTypeCompatibility() As Byte
    DataTypeCompatibility = mDataTypeCompatibility
End Property

Public Property Let DataTypeCompatibility(ByVal Value As Byte)
    mDataTypeCompatibility = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get DataSource() As String
    DataSource = mDataSource
End Property

Public Property Let DataSource(ByVal Value As String)
    mDataSource = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get FailoverPartner() As String
    FailoverPartner = mFailoverPartner
End Property

Public Property Let FailoverPartner(ByVal Value As String)
    mFailoverPartner = Value
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

Public Property Get Backup_RemoteFolder() As String
    Backup_RemoteFolder = mBackup_RemoteFolder
End Property

Public Property Let Backup_RemoteFolder(ByVal Value As String)
    mBackup_RemoteFolder = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get Backup_LocalFolder() As String
    Backup_LocalFolder = mBackup_LocalFolder
End Property

Public Property Let Backup_LocalFolder(ByVal Value As String)
    mBackup_LocalFolder = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get Backup_FileName() As String
    Backup_FileName = mBackup_FileName
End Property

Public Property Let Backup_FileName(ByVal Value As String)
    mBackup_FileName = Value
End Property

'///////////////////////////////////////////////////////

Public Property Get ParametersSaveToRegistry() As Boolean
    ParametersSaveToRegistry = mParametersSaveToRegistry
End Property

Public Property Let ParametersSaveToRegistry(ByVal Value As Boolean)
    mParametersSaveToRegistry = Value
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
    
    mProvider = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "Provider", "", csrdtString)
    mConnectionTimeout = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "ConnectionTimeout", 15, csrdtNumberInteger)
    mCommandTimeout = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "CommandTimeout", 15, csrdtNumberInteger)
    mPacketSize = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "PacketSize", 4096, csrdtNumberInteger)
    mDataTypeCompatibility = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "DataTypeCompatibility", 80, csrdtNumberInteger)
    mDataSource = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "DataSource", "", csrdtString)
    mFailoverPartner = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "FailoverPartner", "", csrdtString)
    mUserID = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "UserID", "", csrdtString)
    mPassword = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "Password", "", csrdtString)
    If mPassword <> "" Then
        Set DES = New CSC_Encryption_DES
        mPassword = DES.DecryptString(mPassword, DES.PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    mDatabase = CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "Name", "", csrdtString)
    
    mReportsPath = CSM_Registry.GetValue_FromApplication_LocalMachine("Reports", "Path", "", csrdtString)
    
    mBackup_RemoteFolder = CSM_Registry.GetValue_FromApplication_LocalMachine("RemoteFolder", "Backup", "", csrdtString)
    mBackup_LocalFolder = CSM_Registry.GetValue_FromApplication_LocalMachine("LocalFolder", "Backup", "", csrdtString)
    mBackup_FileName = CSM_Registry.GetValue_FromApplication_LocalMachine("FileName", "Backup", "", csrdtString)
    
    LoadParametersFromRegistry = True
End Function

Public Function LoadParametersFromINI() As Boolean
    Dim DES As CSC_Encryption_DES
    
    mProvider = CSM_INI.GetValue_FromApplication("Database", "Provider", "", csidtString)
    mConnectionTimeout = CSM_INI.GetValue_FromApplication("Database", "ConnectionTimeout", 15, csidtNumberInteger)
    mCommandTimeout = CSM_INI.GetValue_FromApplication("Database", "CommandTimeout", 15, csidtNumberInteger)
    mPacketSize = CSM_INI.GetValue_FromApplication("Database", "PacketSize", 4096, csidtNumberInteger)
    mDataTypeCompatibility = CSM_INI.GetValue_FromApplication("Database", "DataTypeCompatibility", 80, csidtNumberInteger)
    mDataSource = CSM_INI.GetValue_FromApplication("Database", "DataSource", "", csidtString)
    mFailoverPartner = CSM_INI.GetValue_FromApplication("Database", "FailoverPartner", "", csidtString)
    mUserID = CSM_INI.GetValue_FromApplication("Database", "UserID", "", csidtString)
    mPassword = CSM_INI.GetValue_FromApplication("Database", "Password", "", csidtString)
    If mPassword <> "" Then
        Set DES = New CSC_Encryption_DES
        mPassword = DES.DecryptString(mPassword, DES.PASSWORD_ENCRYPTION_KEY)
        Set DES = Nothing
    End If
    mDatabase = CSM_INI.GetValue_FromApplication("Database", "Name", "", csidtString)
    
    mReportsPath = CSM_INI.GetValue_FromApplication("Reports", "Path", "", csidtString)
    
    mBackup_RemoteFolder = CSM_INI.GetValue_FromApplication("Backup", "RemoteFolder", "", csidtString)
    mBackup_LocalFolder = CSM_INI.GetValue_FromApplication("Backup", "LocalFolder", "", csidtString)
    mBackup_FileName = CSM_INI.GetValue_FromApplication("Backup", "FileName", "", csidtString)

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
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "PacketSize", mPacketSize)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "DataTypeCompatibility", mDataTypeCompatibility)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "DataSource", mDataSource)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "FailoverPartner", mFailoverPartner)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "UserID", mUserID)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "Password", mPassword)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "Name", mDatabase)
    
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Reports", "Path", mReportsPath)
    
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Backup", "RemoteFolder", mBackup_RemoteFolder)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Backup", "LocalFolder", mBackup_LocalFolder)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Backup", "FileName", mBackup_FileName)
    
    SaveParametersToRegistry = True
End Function

Public Function SaveParametersToINI() As Boolean
    Call CSM_INI.SetValue_ToApplication("Database", "Provider", mProvider)
    Call CSM_INI.SetValue_ToApplication("Database", "ConnectionTimeout", mConnectionTimeout)
    Call CSM_INI.SetValue_ToApplication("Database", "CommandTimeout", mCommandTimeout)
    Call CSM_INI.SetValue_ToApplication("Database", "PacketSize", mPacketSize)
    Call CSM_INI.SetValue_ToApplication("Database", "DataTypeCompatibility", mDataTypeCompatibility)
    Call CSM_INI.SetValue_ToApplication("Database", "DataSource", mDataSource)
    Call CSM_INI.SetValue_ToApplication("Database", "FailoverPartner", mFailoverPartner)
    Call CSM_INI.SetValue_ToApplication("Database", "UserID", mUserID)
    Call CSM_INI.SetValue_ToApplication("Database", "Password", mPassword)
    Call CSM_INI.SetValue_ToApplication("Database", "Name", mDatabase)
    
    Call CSM_INI.SetValue_ToApplication("Reports", "Path", mReportsPath)
    
    Call CSM_INI.SetValue_ToApplication("Backup", "RemoteFolder", mBackup_RemoteFolder)
    Call CSM_INI.SetValue_ToApplication("Backup", "LocalFolder", mBackup_LocalFolder)
    Call CSM_INI.SetValue_ToApplication("Backup", "FileName", mBackup_FileName)
    
    SaveParametersToINI = True
End Function

Public Property Get Connection() As ADODB.Connection
    Set Connection = mConnection
End Property

Public Function Connect() As Long

    Screen.MousePointer = vbHourglass
    
    CSM_ApplicationLog.WriteLogEvent "Database Connection: INIT", vbLogEventTypeInformation, pParametro.LogAccion_Enabled

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Set mConnection = New ADODB.Connection
    
    With mConnection
        .Provider = mProvider
        .ConnectionTimeout = mConnectionTimeout
        .CommandTimeout = mCommandTimeout
        .CursorLocation = IIf(mCursorLocationServer, adUseServer, adUseClient)
        .ConnectionString = "Data Source=" & mDataSource & IIf(mFailoverPartner <> "", ";FailoverPartner=" & mFailoverPartner, "") & ";User ID=" & mUserID & ";Password=" & mPassword & ";Packet Size=" & mPacketSize & ";DataTypeCompatibility=" & mDataTypeCompatibility & ";MARS Connection=True"
        .Open
        If mDatabase <> "" Then
            .DefaultDatabase = mDatabase
        End If
    End With
    
    Connect = -1
    CSM_ApplicationLog.WriteLogEvent "Database Connection: DONE", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    If mConnection.Errors.Count > 0 Then
        If mConnection.Errors(0).NativeError = ERROR_DATABASE_NOTEXIST Then
            CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_SQL.Connect", "La Base de Datos especificada no existe." & vbCr & vbCr & "Data Source: " & mDataSource & vbCr & "Database: " & mDatabase
        Else
            CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_SQL.Connect", "Error al crear la conexión al Servidor de Base de Datos." & vbCr & vbCr & "Data Source: " & mDataSource
        End If
    Else
        CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_SQL.Connect", "Error al crear la conexión al Servidor de Base de Datos." & vbCr & vbCr & "Data Source: " & mDataSource
    End If
End Function

Public Function LoadParametersFromDatabase() As Boolean
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
        WriteLogEvent "Closing Database Connection", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
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

Public Function OpenRecordset(ByRef Recordset As ADODB.Recordset, ByVal SQLStatement As String, ByVal CursorType As CursorTypeEnum, ByVal LockType As LockTypeEnum, ByVal Options As Long, ByVal errorMessage As String, ByVal ErrorSource As String) As Boolean
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
    CSM_Error.ShowErrorMessage ErrorSource, errorMessage
    On Error Resume Next
    ADODBRecordset.Close
    Set ADODBRecordset = Nothing
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetFirstValueFromRecordset(ByVal SQLStatement As String, ByVal errorMessage As String, ByVal ErrorSource As String) As Variant
    Dim ADODBRecordset As ADODB.Recordset
    
    If OpenRecordset(ADODBRecordset, SQLStatement, adOpenForwardOnly, adLockReadOnly, adCmdText, errorMessage, ErrorSource) Then
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
    CSM_Error.ShowErrorMessage ErrorSource, errorMessage
    On Error Resume Next
    ADODBRecordset.Close
    Set ADODBRecordset = Nothing
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetDataTypeArrayFromRecordset(ByRef recData As ADODB.Recordset) As Variant
    Dim Index As Integer
    Dim aDataTypes() As ADODB.DataTypeEnum
    
    If Not recData Is Nothing Then
        ReDim aDataTypes(0 To recData.Fields.Count - 1)
        For Index = 0 To recData.Fields.Count - 1
            aDataTypes(Index) = recData(Index).Type
        Next Index
    End If
    GetDataTypeArrayFromRecordset = aDataTypes()
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function WriteBLOB(ByVal SourceFileName As String, ByRef Recordset As Recordset, ByVal FieldName As String) As Boolean
    Dim BlocksCount As Integer
    Dim BlockNumber As Integer
    Dim SourceFileNumber As Integer
    Dim FileLength As Long
    Dim LeftOver As Long
    Dim FileData As String

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    'Open the source file.
    SourceFileNumber = FreeFile
    Open SourceFileName For Binary Access Read As SourceFileNumber

    'Get the length of the file.
    FileLength = LOF(SourceFileNumber)
    If FileLength = 0 Then
        Recordset(FieldName).AppendChunk Null
        Recordset.Update
        WriteBLOB = True
        Exit Function
    End If

    'Calculate the number of blocks to read and leftover bytes.
    BlocksCount = FileLength \ BLOB_BLOCK_SIZE
    LeftOver = FileLength Mod BLOB_BLOCK_SIZE

    'Read the remaining blocks of data, writing them to the table.
    FileData = String(BLOB_BLOCK_SIZE, 32)
    For BlockNumber = 1 To BlocksCount
        Get SourceFileNumber, , FileData
        Recordset(FieldName).AppendChunk (FileData)
    Next BlockNumber

    'Read the leftover data, writing it to the table.
    If LeftOver > 0 Then
        FileData = String(LeftOver, 32)
        Get SourceFileNumber, , FileData
        Recordset(FieldName).AppendChunk (FileData)
    End If

    Close SourceFileNumber
    WriteBLOB = True
    Exit Function

ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_SQL.WriteBLOB", "Error al guardar los datos BLOB en la Base de Datos."
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function ReadBLOB(ByRef Recordset As ADODB.Recordset, ByVal FieldName As String, ByVal DestinationFileName As String) As Boolean
    Dim BlocksCount As Integer
    Dim BlockNumber As Integer
    Dim DestinationFileNumber As Integer
    Dim FileLength As Long
    Dim LeftOver As Long
    Dim FileData As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    'Get the size of the field.
    FileLength = Recordset(FieldName).ActualSize()
    If FileLength = 0 Then
        Exit Function
    End If
    
    'Calculate number of blocks to write and leftover bytes.
    BlocksCount = FileLength \ BLOB_BLOCK_SIZE
    LeftOver = FileLength Mod BLOB_BLOCK_SIZE
    
    'Remove any existing destination file.
    DestinationFileNumber = FreeFile
    Open DestinationFileName For Output As DestinationFileNumber
    Close DestinationFileNumber
    
    'Open the destination file.
    Open DestinationFileName For Binary As DestinationFileNumber
    
    'Write the blocks of data to the output file.
    For BlockNumber = 1 To BlocksCount
        'Reads a chunk and writes it to output file.
        FileData = Recordset(FieldName).GetChunk(BLOB_BLOCK_SIZE)
        Put DestinationFileNumber, , FileData
    Next BlockNumber
    
    'Write the leftover data to the output file.
    If LeftOver > 0 Then
        FileData = Recordset(FieldName).GetChunk(LeftOver)
        Put DestinationFileNumber, , FileData
    End If
    
    Close DestinationFileNumber
    ReadBLOB = FileLength
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_SQL.ReadBLOB", "Error al leer los datos BLOB de la Base de Datos."
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function ExecuteScript(ByVal script As String, ByVal errorMessage As String, Optional ByVal getResultRecordset As Boolean = False, Optional ByRef recResult As ADODB.Recordset = Nothing, Optional Database As String = "", Optional TIMEOUT As Long = 0) As Boolean
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    If Database <> "" Then
        pDatabase.Connection.DefaultDatabase = "[" & Database & "]"
    End If
    pDatabase.Connection.CommandTimeout = TIMEOUT
    If getResultRecordset Then
        Set recResult = pDatabase.Connection.Execute(script)
    Else
        pDatabase.Connection.Execute script
    End If
    pDatabase.Connection.CommandTimeout = pParametro.Database_CommandTimeout
    If Database <> "" Then
        pDatabase.Connection.DefaultDatabase = pParametro.Database_Database
    End If

    Screen.MousePointer = vbDefault
    ExecuteScript = True
    Exit Function

ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_SQL.ExecuteScript", errorMessage & vbCr & vbCr & "Script SQL:" & vbCr & script
    pDatabase.Connection.CommandTimeout = pParametro.Database_CommandTimeout
    If Database <> "" Then
        pDatabase.Connection.DefaultDatabase = pParametro.Database_Database
    End If
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function GetPrimaryFile_PathAndFileName() As String
    Dim cmdFiles As ADODB.command
    Dim recFiles As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
        
    Set cmdFiles = New ADODB.command
    With cmdFiles
        Set .ActiveConnection = pDatabase.Connection
        .CommandType = adCmdStoredProc
        .CommandText = "sp_helpfile"
    End With
    
    Set recFiles = New ADODB.Recordset
    With recFiles
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open cmdFiles
        
        If Not .EOF Then
            GetPrimaryFile_PathAndFileName = .Fields("filename").Value
        End If
        
        .Close
        Set recFiles = Nothing
    End With
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.CSC_Database_ADO_SQL.GetPrimaryFile_PathAndFileName", "Error al obtener el Archivo de Datos Primario de la Base de Datos."
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Sub ShowUpdateErrorMessage(ByVal Source As String, ByVal UnknownErrorMessage As String, ParamArray DuplicateKeyMessagePair())
    Dim KeyName As Variant
    Dim Index As Integer
    
    If UBound(DuplicateKeyMessagePair) > 0 And Err.Number = ERROR_DUPLICATE_PRIMARYKEY Or Err.Number = ERROR_DUPLICATEKEY_UNIQUEINDEX Then
        For Index = 0 To UBound(DuplicateKeyMessagePair) Step 2
            If DuplicateKeyMessagePair(Index) = "" Or DuplicateKeyMessagePair(Index) = "*" Then
                If Index + 1 <= UBound(DuplicateKeyMessagePair) Then
                    Screen.MousePointer = vbDefault
                    MsgBox DuplicateKeyMessagePair(Index + 1), vbExclamation, App.Title
                    Exit Sub
                Else
                    CSM_Error.ShowErrorMessage Source, UnknownErrorMessage
                    Exit Sub
                End If
            ElseIf InStr(1, Err.Description, DuplicateKeyMessagePair(Index)) > 0 Then
                If Index + 1 <= UBound(DuplicateKeyMessagePair) Then
                    Screen.MousePointer = vbDefault
                    MsgBox DuplicateKeyMessagePair(Index + 1), vbExclamation, App.Title
                    Exit Sub
                Else
                    CSM_Error.ShowErrorMessage Source, UnknownErrorMessage
                    Exit Sub
                End If
            End If
        Next Index
        CSM_Error.ShowErrorMessage Source, UnknownErrorMessage
    Else
        CSM_Error.ShowErrorMessage Source, UnknownErrorMessage
    End If
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Sub ShowDeleteErrorMessage(ByVal Source As String, ByVal UnknownErrorMessage As String, ByVal RelatedErrorMessage As String)
    Dim KeyName As Variant
    Dim Index As Integer
    
    Select Case Err.Number
        Case ERROR_RELATED_RECORDS
            Screen.MousePointer = vbDefault
            MsgBox RelatedErrorMessage, vbExclamation, App.Title
        Case Else
            CSM_Error.ShowErrorMessage Source, UnknownErrorMessage
    End Select
End Sub

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub Class_Initialize()
    mParametersSaveToRegistry = True
    Set mConnection = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If mConnection.State <> adStateClosed Then
        mConnection.Close
    End If
    Set mConnection = Nothing
End Sub

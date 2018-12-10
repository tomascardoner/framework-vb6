VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_Database_SQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFilePath As String
Private mFileName As String
Private mFilePathAndName As String

Private mConnection As vbRichClient5.cConnection

Public Function Connect() As Long
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If

    Set mConnection = New vbRichClient5.cConnection
    
    With mConnection
        .OpenDB mFilePathAndName
    End With
    
    Connect = -1
    CSM_ApplicationLog.WriteLogEvent "Database Connection: DONE", vbLogEventTypeInformation
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Classes.CSC_Database_SQLite.Connect", "Error al crear la conexión a la Base de Datos SQLite." & vbCr & vbCr & "File Name: " & mFilePathAndName
End Function

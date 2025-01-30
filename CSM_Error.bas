Attribute VB_Name = "CSM_Error"
Option Explicit

Public Const ERROR_ELEMENT_NOT_FOUND As Long = 35601
Public Const ERROR_TYPE_MISMATCH As Long = 13

Public Sub ShowErrorMessage(ByVal vstrSource As String, Optional ByVal vstrMessage As String, Optional ByVal vstrHelpFile As String, Optional ByVal vlngContext As Long, Optional ByVal vblnShowMessageBox As Boolean = True)
    Dim strPrompt As String
    Dim strToLog As String
    
    Screen.MousePointer = vbDefault
        
    If vblnShowMessageBox Then
        MsgBox vstrMessage & vbCrLf & vbCrLf & "Error: " & Err.Number & " - " & Err.Description, vbOKOnly + vbExclamation, App.Title
    Else
        strPrompt = "Se ha encontrado un Error inesperado."
        strPrompt = strPrompt & vbCr & "Anote la siguiente informaci�n e informela al servicio t�cnico."
        
        strPrompt = strPrompt & vbCr & vbCr & "Origen: " & vstrSource
        strToLog = "Where: " & vstrSource
        
        If vstrMessage <> "" Then
            strPrompt = strPrompt & vbCr & vbCr & vstrMessage
            strToLog = strToLog & " // User Message: " & Replace(vstrMessage, vbCr, " � ")
        End If
        
        strPrompt = strPrompt & vbCr & vbCr & "Error " & Err.Number & ": " & Err.Description & vbCr & vbCr & Err.Source
        strToLog = strToLog & " // VB Error: " & Err.Number & " - " & Err.Description & " // Context: " & Err.Source
                
        CSM_ApplicationLog.WriteLogEvent strToLog, vbLogEventTypeError, pLogEnabled
    End If
End Sub

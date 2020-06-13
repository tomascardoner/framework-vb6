Attribute VB_Name = "CSM_ApplicationLog"
Option Explicit

Private mFilePathAndName As String

Public Sub InitLogging(Optional ByVal Path As String = "", Optional ByVal FileNameTemplate As String = "@EXENAME@_", Optional ByVal MonthsToKeep As Integer = 3)
    Dim FileNameTemplateResolved As String
    Dim FileName As String
    
    Dim DirResult As String
    Dim TempName As String
    Dim FileYear As Integer
    Dim FileMonth As Byte
    
    If Not pIsCompiled Then
        Exit Sub
    End If
    
    'PATH
    If Path = "" Then
        Path = App.Path
    End If
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    'RESOLVE FILE NAME
    If FileNameTemplate = "" Then
        FileNameTemplate = "@EXENAME@_"
    End If
    FileNameTemplateResolved = FileNameTemplate
    FileNameTemplateResolved = Replace(FileNameTemplateResolved, "@EXENAME@", App.EXEName)
    FileNameTemplateResolved = Replace(FileNameTemplateResolved, "@USERNAME@", CSM_Session.GetUserName())
    FileName = FileNameTemplateResolved & Year(Date) & "-" & Format(Month(Date), "00") & ".log"
    
    mFilePathAndName = Path & FileName
    
    
    Call CreateFileHeader
    
    'Clean up log files older than MonthsToKeep months
    DirResult = FileSystem.Dir(Path & FileNameTemplateResolved & "*.log")
    Do While DirResult <> ""
        TempName = Left(Right(DirResult, 11), 7)
        If IsNumeric(Left(TempName, 4)) Then
            FileYear = CInt(Left(TempName, 4))
            If FileYear >= 2000 And FileYear <= 2999 Then
                If IsNumeric(Right(TempName, 2)) Then
                    FileMonth = CByte(Right(TempName, 2))
                    
                    If DateDiff("m", DateSerial(FileYear, FileMonth, 1), Date) > MonthsToKeep Then
                        On Error Resume Next
                        FileSystem.Kill Path & DirResult
                        If Err.Number <> 0 Then
                            'MsgBox "Error Deleting File: " & strDirResult, vbExclamation, App.Title
                        Else
                            WriteLogEvent "Old Log File Deleted: " & Path & DirResult, vbLogEventTypeInformation, True
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
        DirResult = Dir()
    Loop
End Sub

Private Sub CreateFileHeader()
    Dim FileNumber As Integer
    
    FileNumber = FreeFile()
    Open mFilePathAndName For Append As #FileNumber
    
    Print #FileNumber, ""
    Print #FileNumber, ""
    Print #FileNumber, String(80, "=")
    Print #FileNumber, String(80, "=")
    Print #FileNumber, String(80, "=")
    Print #FileNumber, "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    Print #FileNumber, "Started from: " & App.Path & "\" & App.EXEName & ".exe"
    Print #FileNumber, String(80, "=")
    Print #FileNumber, "Type # Date/Time # Description"
    Print #FileNumber, "------------------------------"
    
    Close #FileNumber
End Sub

Public Sub WriteLogEvent(ByVal Description As String, ByVal LogEventType As LogEventTypeConstants, ByVal LogEnabled As Boolean)
    Dim FileNumber As Integer
    Dim LogEventTypeSymbol As String
    
    If Not (pIsCompiled And LogEnabled) Then
        Exit Sub
    End If
    
    FileNumber = FreeFile()
    Open mFilePathAndName For Append As #FileNumber
    
    Select Case LogEventType
        Case vbLogEventTypeError
            LogEventTypeSymbol = "ERR"
        Case vbLogEventTypeWarning
            LogEventTypeSymbol = "WRN"
        Case vbLogEventTypeInformation
            LogEventTypeSymbol = "INF"
    End Select
    
    Print #FileNumber, LogEventTypeSymbol & " # " & Now & " # " & Description
    
    Close #FileNumber
End Sub

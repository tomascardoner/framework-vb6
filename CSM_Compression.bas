Attribute VB_Name = "CSM_Compression"
Option Explicit

'//source was in C# from urls:
'//http://www.codeproject.com/csharp/CompressWithWinShellAPICS.asp
'//http://www.codeproject.com/csharp/DecompressWinShellAPICS.asp

'//set reference to "Microsoft Shell Controls and Automation"

'http://forums.microsoft.com/MSDN/ShowPost.aspx?PostID=1090552&SiteID=1
'Be aware when using the shell automation interface to unzip files as it
'leaves copies of the zip files in the temp directory (defined by %TEMP%).
'Folders named "Temporary Directory X for demo.zip" are generated where X
'is a sequential number from 1 - 99.  When it reaches 99 you will then get
'a error dialog saying "The file exists" and it will not continue.
'I 've no idea why Windows doesn't clean up after itself when unzipping files,
'but it is most annoying...

'//CopyHere options
'0 Default. No options specified.
'4 Do not display a progress dialog box.
'8 Rename the target file if a file exists at the target location with the same name.
'16 Click "Yes to All" in any dialog box displayed.
'64 Preserve undo information, if possible.
'128 Perform the operation only if a wildcard file name (*.*) is specified.
'256 Display a progress dialog box but do not show the file names.
'512 Do not confirm the creation of a new directory if the operation requires one to be created.
'1024 Do not display a user interface if an error occurs.
'4096 Disable recursion.
'9182 Do not copy connected files as a group. Only copy the specified files.

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Zip_Activity(ByVal Action As String, ByVal sFileSource As String, ByVal sFileDest As String)
    '//copies contents of folder to zip file
'    Dim ShellClass  As shell32.Shell
'    Dim Filesource  As shell32.Folder
'    Dim Filedest    As shell32.Folder
'    Dim Folderitems As shell32.Folderitems
'
'    If sFileSource = "" Or sFileDest = "" Then
'        Exit Sub
'    End If
'
'    On Error GoTo ErrorHandler
'
'    Select Case UCase$(Action)
'        Case "ZIPFILE"
'
'            If Right$(UCase$(sFileDest), 4) <> ".ZIP" Then
'                sFileDest = sFileDest & ".ZIP"
'            End If
'
'            If Not Create_Empty_Zip(sFileDest) Then
'                GoTo ErrorHandler
'            End If
'
'            Set ShellClass = New shell32.Shell
'            Set Filedest = ShellClass.Namespace(sFileDest)
'
'            Call Filedest.CopyHere(sFileSource, 20)
'
'        Case "ZIPFOLDER"
'
'            If Right$(UCase$(sFileDest), 4) <> ".ZIP" Then
'                sFileDest = sFileDest & ".ZIP"
'            End If
'
'            If Not Create_Empty_Zip(sFileDest) Then
'                GoTo ErrorHandler
'            End If
'
'            '//Copy a folder and its contents into the newly created zip file
'            Set ShellClass = New shell32.Shell
'            Set Filesource = ShellClass.Namespace(sFileSource)
'            Set Filedest = ShellClass.Namespace(sFileDest)
'            Set Folderitems = Filesource.items
'
'            Call Filedest.CopyHere(Folderitems, 20)
'
'        Case "UNZIP"
'
'            If Right$(UCase$(sFileSource), 4) <> ".ZIP" Then
'                sFileSource = sFileSource & ".ZIP"
'            End If
'
'            Set ShellClass = New shell32.Shell
'            Set Filesource = ShellClass.Namespace(sFileSource)      '//should be zip file
'            Set Filedest = ShellClass.Namespace(sFileDest)          '//should be directory
'            Set Folderitems = Filesource.items                      '//copy zipped items to directory
'
'            Call Filedest.CopyHere(Folderitems, 20)
'
'        Case Else
'
'    End Select
'
'    '//Ziping a file using the Windows Shell API creates another thread where the zipping is executed.
'    '//This means that it is possible that this console app would end before the zipping thread
'    '//starts to execute which would cause the zip to never occur and you will end up with just
'    '//an empty zip file. So wait a second and give the zipping thread time to get started.
'
'    Call Sleep(1000)
'
'ErrorHandler:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description, vbExclamation, "error"
'    End If
'
'    Set ShellClass = Nothing
'    Set Filesource = Nothing
'    Set Filedest = Nothing
'    Set Folderitems = Nothing
End Sub

Private Function Create_Empty_Zip(ByVal sFileName As String) As Boolean
    Dim EmptyZip() As Byte
    Dim j As Integer

    On Error GoTo ErrorHandler
    
    Create_Empty_Zip = False

    '//create zip header
    ReDim EmptyZip(1 To 22)

    EmptyZip(1) = 80
    EmptyZip(2) = 75
    EmptyZip(3) = 5
    EmptyZip(4) = 6
    
    For j = 5 To UBound(EmptyZip)
        EmptyZip(j) = 0
    Next

    '//create empty zip file with header
    Open sFileName For Binary Access Write As #1

    For j = LBound(EmptyZip) To UBound(EmptyZip)
        Put #1, , EmptyZip(j)
    Next
    
    Close #1
    
    Create_Empty_Zip = True

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "Error"
    End If
End Function

Public Function SingleFileTopZIP(ByVal SourceFile As String, ByVal DestinationFile As String) As Boolean
    Dim oFSO As Object
    Dim oApp As Object
    Dim arrHex()
    Dim sBin As String
    Dim i As Byte
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    arrHex = Array(80, 75, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

    For i = 0 To UBound(arrHex)
        sBin = sBin & Chr(arrHex(i))
    Next i
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    With oFSO.CreateTextFile(DestinationFile, True)
        .Write sBin
        .Close
    End With
    Set oFSO = Nothing
    
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(CStr(DestinationFile)).CopyHere (SourceFile)
    Set oApp = Nothing
    
    SingleFileTopZIP = True
    Exit Function
    
ErrorHandler:
    Call CSM_Error.ShowErrorMessage("Forms.frmMDI.SingleFileTopZIP", "Error al comprimir los archivos.")
    On Error Resume Next
    If FileSystem.Dir(DestinationFile) <> "" Then
        Call FileSystem.Kill(DestinationFile)
    End If
End Function

Public Function ExtractFilesFromZIP(ByVal ZIPFileName As String, ByVal DestinationFolder As String) As Boolean
    Dim oApp As Object

    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If ZIPFileName = "" Or DestinationFolder = "" Then
        Exit Function
    End If
    If Right(DestinationFolder, 1) <> "\" Then
        DestinationFolder = DestinationFolder & "\"
    End If

    'Extract the files into the newly created folder
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(CStr(DestinationFolder)).CopyHere oApp.Namespace(CStr(ZIPFileName)).items
    Set oApp = Nothing

    'If you want to extract only one file you can use this:
    'oApp.Namespace(FileNameFolder).CopyHere _
     'oApp.Namespace(Fname).items.Item("test.txt")

    ExtractFilesFromZIP = True
    Exit Function
    
ErrorHandler:
    Call CSM_Error.ShowErrorMessage("Forms.frmMDI.ExtractFilesFromZIP", "Error al descomprimir el archivo.")
End Function

Public Function Zip7_Compress(ByVal PathAndFileName As String) As Boolean
    If pIsCompiled Then
        On Error GoTo ErrorHandler
    End If
    
    CSM_Instance.Execute "7z.exe a -t7z " & CSM_File.GetFileNameWithoutExtension(PathAndFileName) & ".7z " & CSM_File.GetFileName(PathAndFileName) & " -mx=9", , , """" & CSM_File.GetPath(PathAndFileName) & """"
    
    Exit Function
    
ErrorHandler:
End Function

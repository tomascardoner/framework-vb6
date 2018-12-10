Attribute VB_Name = "CSM_Installation"
Option Explicit

Public Function CreateShortcut(ByVal ShortcutPathAndFileName As String, ByVal TargetPath As String, ByVal Arguments As String, ByVal WorkingDirectory As String, ByVal Description As String, ByVal HotKey As String, ByVal IconLocation As String, ByVal RelativePath As String, ByVal WindowStyle As Long) As Boolean
    Dim oShell As IWshRuntimeLibrary.WshShell
    Dim oShortcut As IWshRuntimeLibrary.WshShortcut
    Dim oURLShortcut As IWshRuntimeLibrary.WshURLShortcut
    
    If ShortcutPathAndFileName = "" Then
        CreateShortcut = False
    Else
        On Error GoTo ErrorHandler
        
        Select Case CSM_File.GetFileExtension(ShortcutPathAndFileName)
            Case "lnk"
                Set oShell = New IWshRuntimeLibrary.WshShell
                Set oShortcut = oShell.CreateShortcut(ShortcutPathAndFileName)
                With oShortcut
                    .TargetPath = TargetPath
                    .Arguments = Arguments
                    .WorkingDirectory = WorkingDirectory
                    .Description = Description
                    .HotKey = HotKey
                    If IconLocation <> "" Then
                        .IconLocation = IconLocation
                    End If
                    .RelativePath = RelativePath
                    .WindowStyle = WindowStyle
                
                    .Save
                End With
                Set oShell = Nothing
                CreateShortcut = True
                
            Case "url"
                Set oShell = New IWshRuntimeLibrary.WshShell
                Set oURLShortcut = oShell.CreateShortcut(ShortcutPathAndFileName)
                With oURLShortcut
                    .TargetPath = TargetPath
                    .Arguments = Arguments
                    .WorkingDirectory = WorkingDirectory
                    .Description = Description
                    .HotKey = HotKey
                    If IconLocation <> "" Then
                        .IconLocation = IconLocation
                    End If
                    .RelativePath = RelativePath
                    .WindowStyle = WindowStyle
                
                    .Save
                End With
                Set oShell = Nothing
                CreateShortcut = True
                
            Case Else
                'User input an invalid path or filename
                CreateShortcut = False
        End Select
    End If
    Exit Function
    
ErrorHandler:
End Function

Public Function CreateShortcutDesktop(ByVal ShortcutFileName As String, ByVal TargetPath As String, ByVal Arguments As String, ByVal WorkingDirectory As String, ByVal Description As String, ByVal HotKey As String, ByVal IconLocation As String, ByVal RelativePath As String, ByVal WindowStyle As Long) As Boolean
    Dim oShell As IWshRuntimeLibrary.WshShell
    
    Set oShell = New IWshRuntimeLibrary.WshShell

    CreateShortcutDesktop = CreateShortcut(oShell.SpecialFolders("Desktop") & "\" & ShortcutFileName, TargetPath, Arguments, WorkingDirectory, Description, HotKey, IconLocation, RelativePath, WindowStyle)
    
    Set oShell = Nothing
End Function

Public Function CreateShortcutDesktopAllUsers(ByVal ShortcutFileName As String, ByVal TargetPath As String, ByVal Arguments As String, ByVal WorkingDirectory As String, ByVal Description As String, ByVal HotKey As String, ByVal IconLocation As String, ByVal RelativePath As String, ByVal WindowStyle As Long) As Boolean
    Dim oShell As IWshRuntimeLibrary.WshShell
    
    Set oShell = New IWshRuntimeLibrary.WshShell
    
    CreateShortcutDesktopAllUsers = CreateShortcut(oShell.SpecialFolders("AllUsersDesktop") & "\" & ShortcutFileName, TargetPath, Arguments, WorkingDirectory, Description, HotKey, IconLocation, RelativePath, WindowStyle)
    
    Set oShell = Nothing
End Function


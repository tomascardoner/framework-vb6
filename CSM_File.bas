Attribute VB_Name = "CSM_File"
Option Explicit

Public Function GetFileExtension(ByVal FullPath As String) As String
    Dim aFileNameParts() As String
    
    If Trim(FullPath) <> "" Then
        aFileNameParts = Split(FullPath, ".")
        If UBound(aFileNameParts) > 0 Then
            GetFileExtension = aFileNameParts(UBound(aFileNameParts))
        End If
    End If
End Function

Public Function GetFileName(ByVal FullPath As String) As String
    Dim aFileNameParts() As String
    
    If Trim(FullPath) <> "" Then
        aFileNameParts = Split(FullPath, "\")
        GetFileName = aFileNameParts(UBound(aFileNameParts))
    End If
End Function

Public Function GetFileNameWildcardsResult(ByVal FullPathResult As String, ByVal FullPathWithWildcards As String) As String
    Dim WildcardPosition As Integer
    
    If Trim(FullPathResult) <> "" And Trim(FullPathWithWildcards) <> "" Then
        WildcardPosition = InStr(1, CSM_File.GetFileNameWithoutExtension(FullPathWithWildcards), "?")
        If WildcardPosition = 0 Then
            WildcardPosition = InStr(1, CSM_File.GetFileNameWithoutExtension(FullPathWithWildcards), "*")
        End If
        If WildcardPosition > 0 Then
            GetFileNameWildcardsResult = Mid(FullPathResult, WildcardPosition)
        Else
            GetFileNameWildcardsResult = ""
        End If
    End If
End Function

Public Function GetFileNameWithoutExtension(ByVal FullPath As String) As String
    Dim FILENAME As String
    Dim FileExtensionLenght As Integer
    
    FILENAME = GetFileName(FullPath)
    FileExtensionLenght = Len(GetFileExtension(FullPath))
    
    If FileExtensionLenght > 0 Then
        GetFileNameWithoutExtension = Left(FILENAME, Len(FILENAME) - FileExtensionLenght - 1)
    Else
        GetFileNameWithoutExtension = FILENAME
    End If
End Function

Public Function GetPath(ByVal FullPath As String) As String
    Dim aFileNameParts() As String
    
    If Trim(FullPath) <> "" Then
        aFileNameParts = Split(FullPath, "\")
        If UBound(aFileNameParts) > 0 Then
            GetPath = Left(FullPath, Len(FullPath) - Len(aFileNameParts(UBound(aFileNameParts))))
        End If
    End If
End Function

Public Function GetCollectionOfFiles(ByVal Path As String, ByVal FileNameWithWildcards As String, Optional ByVal IncludePathInCollection As Boolean = False, Optional ByVal IncludeFullFileNameInCollection As Boolean = False, Optional ByVal IncludeFileExtensionInCollection As Boolean = False) As Collection
    Dim FILENAME As String
    Dim FileNameToAdd As String
    
    Set GetCollectionOfFiles = New Collection
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    If Trim(Path) <> "" And Trim(FileNameWithWildcards) <> "" Then
        If Right(Path, 1) <> "\" Then
            Path = Path & "\"
        End If
        FILENAME = FileSystem.Dir(Path & FileNameWithWildcards)
        Do While FILENAME <> ""
            If Not (FileSystem.GetAttr(Path & FILENAME) And vbDirectory) Then
                FileNameToAdd = FILENAME
                If IncludePathInCollection Then
                    FileNameToAdd = Path & FileNameToAdd
                End If
                If Not IncludeFullFileNameInCollection Then
                    FileNameToAdd = CSM_File.GetFileNameWildcardsResult(FileNameToAdd, FileNameWithWildcards)
                End If
                If Not IncludeFileExtensionInCollection Then
                    FileNameToAdd = CSM_File.GetFileNameWithoutExtension(FileNameToAdd)
                End If
                
                GetCollectionOfFiles.Add FileNameToAdd
            End If
            FILENAME = FileSystem.Dir()
        Loop
    End If
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Modules.CSM_File.GetCollectionOfFiles", "Error al obtener la lista de archivos."
End Function

Public Function CopyFilesFromFolder(ByVal RemoteFolder As String, ByVal LocalFolder As String, ByVal FileNameWithWildcards As String)
    Dim CFileName As Collection
    Dim FILENAME As Variant
    
    If Right(RemoteFolder, 1) <> "\" Then
        RemoteFolder = RemoteFolder + "\"
    End If
    If Right(LocalFolder, 1) <> "\" Then
        LocalFolder = LocalFolder + "\"
    End If
    
    Set CFileName = GetCollectionOfFiles(RemoteFolder, FileNameWithWildcards, False, True, True)

    For Each FILENAME In CFileName
        'VERIFICO SI EXISTE EL ARCHIVO EN LA CARPETA LOCAL
        If FileSystem.Dir(LocalFolder & CStr(FILENAME)) = "" Then
            'NO EXISTE, COPIO
            FileSystem.FileCopy RemoteFolder & CStr(FILENAME), LocalFolder & CStr(FILENAME)
        Else
            'EXISTE, COMPARO FECHAS
            If DateDiff("n", FileSystem.FileDateTime(RemoteFolder & CStr(FILENAME)), FileSystem.FileDateTime(LocalFolder & CStr(FILENAME))) <> 0 Then
                'LAS FECHAS SON DIFERENTES, LAS COPIO
                FileSystem.FileCopy RemoteFolder & CStr(FILENAME), LocalFolder & CStr(FILENAME)
            End If
        End If
    Next FILENAME
    CopyFilesFromFolder = True
    Exit Function
    
ErrorHandler:
    Call CSM_Error.ShowErrorMessage("Modules.CSM_File.CopyFilsFromFolder", "Error al copiar el archivo." & vbCr & vbCr & "Archivo remoto: " & RemoteFolder & CStr(FILENAME) & vbCr & "Archivo local: " & LocalFolder & CStr(FILENAME))
End Function

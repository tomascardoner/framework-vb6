Attribute VB_Name = "CSM_Instance"
Option Explicit

'///////////////////////////////////////////////////////////////////////////////////////
'EXECUTE
Private Const SW_HIDE = 0                   'Hides the window and activates another window.
Private Const SW_MAXIMIZE = 3               'Maximizes the specified window.
Private Const SW_MINIMIZE = 6               'Minimizes the specified window and activates the next top-level window in the Z order.
Private Const SW_RESTORE = 9                'Activates and displays the window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when restoring a minimized window.
Private Const SW_SHOW = 5                   'Activates the window and displays it in its current size and position.
Private Const SW_SHOWDEFAULT = 10           'Sets the show state based on the SW_ flag specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application. An application should call ShowWindow with this flag to set the initial show state of its main window.
Private Const SW_SHOWMAXIMIZED = 3          'Activates the window and displays it as a maximized window.
Private Const SW_SHOWMINIMIZED = 2          'Activates the window and displays it as a minimized window.
Private Const SW_SHOWMINNOACTIVE = 7        'Displays the window as a minimized window. The active window remains active.
Private Const SW_SHOWNA = 8                 'Displays the window in its current state. The active window remains active.
Private Const SW_SHOWNOACTIVATE = 4         'Displays a window in its most recent size and position. The active window remains active.
Private Const SW_SHOWNORMAL = 1             'Activates and displays a window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when displaying the window for the first time.

Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                ' file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                ' path not found
Private Const SE_ERR_OOM = 8                ' out of memory
Private Const SE_ERR_SHARE = 26

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'///////////////////////////////////////////////////////////////////////////////////////
'APPLICATION INSTANCE
Private Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const GW_HWNDPREV = 3

Public Function Execute(ByVal command As String, Optional ByVal eShowCmd As Long = SW_SHOWDEFAULT, Optional ByVal Parameters As String = "", Optional ByVal DefaultDir As String = "", Optional Operation As String = "open", Optional Owner As Long = 0) As String
    Dim lReturn As Long
    Dim sError As String
    
    On Error Resume Next
    
    lReturn = ShellExecute(Owner, Operation, command, Parameters, DefaultDir, eShowCmd)
    If (lReturn < 0) Or (lReturn > 32) Then
        'NO ERROR
        Execute = ""
    Else
        'RETURN ERROR MESSAGE
        Select Case lReturn
            Case 0
                sError = "No hay memoria suficiente"
            Case ERROR_FILE_NOT_FOUND
                sError = "Archivo no encontrado"
            Case ERROR_PATH_NOT_FOUND
                sError = "Ruta no encontrada"
            Case ERROR_BAD_FORMAT
                sError = "El archivo ejecutable no es válido o está corrupto"
            Case SE_ERR_ACCESSDENIED
                sError = "Error de acceso a la Ruta/Archivo"
            Case SE_ERR_ASSOCINCOMPLETE
                sError = "Este tipo de archivo no tiene una asociación válida."
            Case SE_ERR_DDEBUSY
                sError = "El archivo no se puede abrir porque la aplicación de destino está ocupada. Por favor, reinténtelo en un momento."
            Case SE_ERR_DDEFAIL
                sError = "El archivo no se puede abrir porque la transacción de DDE falló. Por favor, reinténtelo en un momento."
            Case SE_ERR_DDETIMEOUT
                sError = "El archivo no se puede abrir porque se agotó el tiempo de espera. Por favor, reinténtelo en un momento."
            Case SE_ERR_DLLNOTFOUND
                sError = "No se encontró la librería de enlace dinámico (DLL)."
            Case SE_ERR_FNF
                sError = "No se encontró el archivo"
            Case SE_ERR_NOASSOC
                sError = "No hay ninguna aplicación asociada a este tipo de archivo."
            Case SE_ERR_OOM
                sError = "No hay memoria suficiente"
            Case SE_ERR_PNF
                sError = "Ruta no encontrada"
            Case SE_ERR_SHARE
                sError = "Ocurrió una violación de compartir."
            Case Else
                sError = "Ocurrió un error mientras se intentaba abrir el archivo."
        End Select
        
        Execute = sError
    End If
End Function

Public Function ExecuteScript(ByVal command As String)
    Dim oShell As Object 'CreateObject("IWshRuntimeLibrary.WshShell")
    Dim lResult As Long
    
    On Error GoTo ErrorHandler
    
    Set oShell = CreateObject("IWshRuntimeLibrary.WshShell") ' New IWshRuntimeLibrary.WshShell
    
    lResult = oShell.Run(command, 1, 0)
    Exit Function
    
ErrorHandler:
End Function


' *****************************************************************************
' Purpose:  Check If Running in Visual Basic IDE or Compiled
'
' Method:   Debug.Print to divide one by zero
'
' Inputs:
'       None
'
' Outputs:
'       Returns True if is Compiled.
'
' Errors:
'       This Function no raise Errors.
'
' Asserts:
'
' Developer                 Date            Comments
' ---------                 ----            --------
' Tomas A. Cardoner                         Initial creation.
' *****************************************************************************
Public Function IsCompiled() As Boolean
    On Error GoTo NotCompiled

    Debug.Print 1 / 0
    IsCompiled = True

NotCompiled:
End Function

Public Sub ActivatePrevious()
    Dim PrevHndl As Long
    Dim Result As Long
    Dim OldTitle As String
    
    'Rename the title of this application so FindWindow
    'will not find this application instance.
    OldTitle = App.Title
    App.Title = "unwanted instance"
    
    PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    
    'Check if found
    If PrevHndl = 0 Then
       'No previous instance found.
       App.Title = OldTitle
       Exit Sub
    End If
    
    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    
    'Restore the program.
    Result = OpenIcon(PrevHndl)
    
    'Activate the application.
    Result = SetForegroundWindow(PrevHndl)
End Sub

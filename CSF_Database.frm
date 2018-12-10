VERSION 5.00
Begin VB.Form CSF_Database 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base de Datos"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CSF_Database.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReportsPath 
      Height          =   315
      Left            =   3360
      TabIndex        =   19
      Top             =   3480
      Width           =   3555
   End
   Begin VB.CommandButton cmdReportsPath 
      Caption         =   "..."
      Height          =   330
      Left            =   6960
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Examinar..."
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txtBackupCopiesNumber 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3360
      TabIndex        =   17
      Tag             =   "INTEGER|EMPTY|NOTZERO|POSITIVE|99"
      Top             =   3060
      Width           =   390
   End
   Begin VB.CommandButton cmdDatabase 
      Caption         =   "Actualizar"
      Height          =   330
      Left            =   6240
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Examinar..."
      Top             =   2640
      Width           =   975
   End
   Begin VB.ComboBox cboDatabase 
      Height          =   330
      Left            =   3360
      TabIndex        =   14
      Top             =   2640
      Width           =   2790
   End
   Begin VB.CommandButton cmdDataSource 
      Caption         =   "..."
      Height          =   330
      Left            =   6960
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Examinar..."
      Top             =   1380
      Width           =   255
   End
   Begin VB.TextBox txtCommandTimeout 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3360
      TabIndex        =   5
      Tag             =   "INTEGER|EMPTY|NOTZERO|POSITIVE|999"
      Top             =   960
      Width           =   510
   End
   Begin VB.TextBox txtConnectionTimeout 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Tag             =   "INTEGER|EMPTY|NOTZERO|POSITIVE|999"
      Top             =   540
      Width           =   510
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   2220
      Width           =   3855
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtDataSource 
      Height          =   315
      Left            =   3360
      TabIndex        =   7
      Top             =   1380
      Width           =   3555
   End
   Begin VB.ComboBox cboProvider 
      Height          =   330
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3870
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   6000
      TabIndex        =   22
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4620
      TabIndex        =   21
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblReportsPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación de los Reportes:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   18
      Top             =   3540
      Width           =   2190
   End
   Begin VB.Label lblBackupCopiesNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copias de Seguridad a Mantener:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   2730
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base de Datos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1005
   End
   Begin VB.Label lblUserID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   1860
      Width           =   675
   End
   Begin VB.Label lblDataSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Origen de los Datos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Label lblCommandTimeout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo de Espera de los Comandos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   3045
   End
   Begin VB.Label lblConnectionTimeout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo de Espera de la Conexión:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2805
   End
   Begin VB.Label lblProvider 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Base de Datos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1875
   End
End
Attribute VB_Name = "CSF_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DUMMY_PASSWORD = "- ¿Creías que era tan fácil? -"

Private Const PROVIDER_MSSQLSERVER_OLEDB = "SQLOLEDB.1"
Private Const PROVIDER_MSSQLSERVER_SQLNATIVE_1 = "SQLNCLI.1"
Private Const PROVIDER_MSSQLSERVER_SQLNATIVE_10 = "SQLNCLI10"
Private Const PROVIDER_MSSQLSERVER_SQLNATIVE_11 = "SQLNCLI11"
Private Const USER_DEFAULT_MSSQLSERVER = "sa"
Private Const PROVIDER_MSACCESS = "Microsoft.Jet.OLEDB.4.0"
Private Const USER_DEFAULT_MSACCESS = "Admin"

Private mPassword As String

Private mKeyDecimal As Boolean

Private Sub Form_Load()
    Caption = App.Title & " - Base de Datos"
    
    Call CSM_Control_TextBox.PrepareAll(Me)
    
    cboProvider.AddItem "Microsoft SQL Server - OLEDB"
    cboProvider.AddItem "Microsoft SQL Server - Native Client 1.0"
    cboProvider.AddItem "Microsoft SQL Server - Native Client 10.0"
    cboProvider.AddItem "Microsoft SQL Server - Native Client 11.0"
    cboProvider.AddItem "Microsoft Access"
    cboProvider.ListIndex = 0

    'INI
    If pDatabase.Provider = PROVIDER_MSSQLSERVER_OLEDB Then
        cboProvider.ListIndex = 0
    ElseIf pDatabase.Provider = PROVIDER_MSSQLSERVER_SQLNATIVE_1 Then
        cboProvider.ListIndex = 1
    ElseIf pDatabase.Provider = PROVIDER_MSSQLSERVER_SQLNATIVE_10 Then
        cboProvider.ListIndex = 2
    ElseIf pDatabase.Provider = PROVIDER_MSSQLSERVER_SQLNATIVE_11 Then
        cboProvider.ListIndex = 3
    ElseIf pDatabase.Provider = PROVIDER_MSACCESS Then
        cboProvider.ListIndex = 4
    Else
        cboProvider.ListIndex = -1
    End If
    txtConnectionTimeout.Text = pDatabase.ConnectionTimeout
    txtCommandTimeout.Text = pDatabase.CommandTimeout
    txtDataSource.Text = pDatabase.DataSource
    txtUserID.Text = pDatabase.UserID
    mPassword = pDatabase.Password
    If mPassword <> "" Then
        txtPassword.Text = DUMMY_PASSWORD
    End If

    cboDatabase.Text = pDatabase.Database
    
    txtBackupCopiesNumber.Text = pDatabase.BackupCopiesNumber
    
    txtReportsPath.Text = pDatabase.ReportsPath
    
    Call CSM_Control_TextBox.FormatAll(Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    mKeyDecimal = CSM_Control_TextBox.CheckKeyDown(ActiveControl, KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(ActiveControl, KeyAscii, mKeyDecimal)
End Sub

Private Sub cboProvider_Click()
    Select Case cboProvider.ListIndex
        Case 0, 1, 2, 3
            txtUserID.Text = USER_DEFAULT_MSSQLSERVER
        Case 4
            txtUserID.Text = USER_DEFAULT_MSACCESS
        Case Else
            txtUserID.Text = ""
    End Select
    txtPassword.Text = ""
    cmdDataSource.Visible = (cboProvider.ListIndex = 4)
    lblDatabase.Visible = (cboProvider.ListIndex <> 4)
    cboDatabase.Visible = (cboProvider.ListIndex <> 4)
    cmdDatabase.Visible = (cboProvider.ListIndex <> 4)
    lblBackupCopiesNumber.Visible = (cboProvider.ListIndex = 4)
    txtBackupCopiesNumber.Visible = (cboProvider.ListIndex = 4)
End Sub

Private Sub cmdDatabase_Click()
    Dim DBConnection As ADODB.Connection
    Dim recTables As ADODB.Recordset
    
    If cboProvider.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Tipo de Base de Datos.", vbInformation, App.Title
        cboProvider.SetFocus
        Exit Sub
    End If
    If Trim(txtDataSource.Text) = "" Then
        MsgBox "Debe especificar el Origen de los Datos.", vbInformation, App.Title
        txtDataSource.SetFocus
        Exit Sub
    End If
    If Trim(txtUserID.Text) = "" Then
        MsgBox "Debe especificar el Usuario.", vbInformation, App.Title
        txtUserID.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    cboDatabase.Clear
    
    Set DBConnection = New ADODB.Connection
    
    DBConnection.Provider = PROVIDER_MSSQLSERVER_OLEDB
    DBConnection.ConnectionTimeout = 15
    DBConnection.CommandTimeout = 30
    DBConnection.CursorLocation = adUseClient
    DBConnection.Mode = adModeShareDenyNone
    DBConnection.ConnectionString = "Data Source=" & txtDataSource.Text
    DBConnection.Open , txtUserID.Text, IIf(txtPassword.Text = DUMMY_PASSWORD, pDatabase.Password, txtPassword.Text)
    DBConnection.DefaultDatabase = "master"
    
    Set recTables = New ADODB.Recordset
    Set recTables.ActiveConnection = DBConnection
    recTables.Source = "SELECT name FROM sysdatabases ORDER BY name"
    recTables.Open , , adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Do While Not recTables.EOF
        cboDatabase.AddItem recTables("name").Value
        recTables.MoveNext
    Loop
    
    recTables.Close
    Set recTables = Nothing
    
    DBConnection.Close
    Set DBConnection = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Forms.CSF_Database.Database_Click", "Error al crear la conexión al Origen de Datos y obtener la lista de Bases de Datos."
End Sub

Private Sub cboProvider_GotFocus()
    CSM_Control_ComboBox.SelAllText cboProvider
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

Private Sub txtConnectionTimeout_GotFocus()
    CSM_Control_TextBox.SelAllText txtConnectionTimeout
End Sub

Private Sub txtConnectionTimeout_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtConnectionTimeout)
End Sub

Private Sub txtCommandTimeout_GotFocus()
    CSM_Control_TextBox.SelAllText txtCommandTimeout
End Sub

Private Sub txtCommandTimeout_LostFocus()
    Call CSM_Control_TextBox.FormatValue_ByTag(txtCommandTimeout)
End Sub

Private Sub txtDataSource_GotFocus()
    CSM_Control_TextBox.SelAllText txtDataSource
End Sub

Private Sub cmdDataSource_Click()
    Dim FileName As String
    
    FileName = CSM_CommonDialog.FileOpen(Me.hwnd, "Seleccione la Base de Datos", "Bases de Datos de Microsoft Access (*.mdb)|*.mdb|Todos los Archivos (*.*)|*.*", IIf(txtDataSource.Text = "", App.Path, txtDataSource.Text))
    If FileName <> "" Then
        txtDataSource.Text = FileName
    End If
    txtDataSource.SetFocus
End Sub

Private Sub txtUserID_GotFocus()
    CSM_Control_TextBox.SelAllText txtUserID
End Sub

Private Sub txtPassword_GotFocus()
    CSM_Control_TextBox.SelAllText txtPassword
End Sub

Private Sub txtBackupCopiesNumber_GotFocus()
    CSM_Control_TextBox.SelAllText txtBackupCopiesNumber
End Sub

Private Sub txtBackupCopiesNumber_LostFocus()
    txtBackupCopiesNumber.Text = Val(txtBackupCopiesNumber.Text)
End Sub

Private Sub txtReportsPath_GotFocus()
    CSM_Control_TextBox.SelAllText txtReportsPath
End Sub

Private Sub cmdReportsPath_Click()
    Dim Path As String
    
    Path = CSM_CommonDialog.BrowseForFolder(Me.hwnd, "Seleccione la ubicación de los Reportes")
    If Path <> "" Then
        txtReportsPath.Text = Path & IIf(Right(Path, 1) = "\", "", "\")
    End If
    txtReportsPath.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim DES As CSC_Encryption_DES
    
    If cboProvider.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Tipo de Base de Datos.", vbInformation, App.Title
        cboProvider.SetFocus
        Exit Sub
    End If
    If Trim(txtDataSource.Text) = "" Then
        MsgBox "Debe especificar el Origen de los Datos.", vbInformation, App.Title
        txtDataSource.SetFocus
        Exit Sub
    End If
    If Trim(txtUserID.Text) = "" Then
        MsgBox "Debe especificar el Usuario.", vbInformation, App.Title
        txtUserID.SetFocus
        Exit Sub
    End If
    If cboDatabase.Visible And Trim(cboDatabase.Text) = "" Then
        MsgBox "Debe especificar la Base de Datos.", vbInformation, App.Title
        cboDatabase.SetFocus
        Exit Sub
    End If
    
    'GET NEW PARAMETERS
    Select Case cboProvider.ListIndex
        Case 0
            pDatabase.Provider = PROVIDER_MSSQLSERVER_OLEDB
        Case 1
            pDatabase.Provider = PROVIDER_MSSQLSERVER_SQLNATIVE_1
        Case 2
            pDatabase.Provider = PROVIDER_MSSQLSERVER_SQLNATIVE_10
        Case 3
            pDatabase.Provider = PROVIDER_MSSQLSERVER_SQLNATIVE_11
        Case 4
            pDatabase.Provider = PROVIDER_MSACCESS
    End Select
    pDatabase.ConnectionTimeout = Val(txtConnectionTimeout.Text)
    pDatabase.CommandTimeout = Val(txtCommandTimeout.Text)
    pDatabase.DataSource = txtDataSource.Text
    pDatabase.UserID = txtUserID.Text
    If txtPassword.Text <> DUMMY_PASSWORD Then
        mPassword = txtPassword.Text
    End If
    If Trim(mPassword) <> "" Then
        Set DES = New CSC_Encryption_DES
        mPassword = DES.EncryptString(mPassword, DES.PASSWORD_ENCRYPTION_KEY, False)
        Set DES = Nothing
    End If
    pDatabase.Password = mPassword
    
    If cboProvider.ListIndex >= 0 And cboProvider.ListIndex <= 3 Then
        pDatabase.Database = cboDatabase.Text
    Else
        pDatabase.Database = ""
    End If
    If cboProvider.ListIndex = 4 Then
        pDatabase.BackupCopiesNumber = Val(txtBackupCopiesNumber.Text)
    Else
        pDatabase.BackupCopiesNumber = 0
    End If
    pDatabase.ReportsPath = txtReportsPath.Text
    
    Call pDatabase.SaveParameters
    
    Me.Tag = "OK"
    Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CSF_Database = Nothing
End Sub

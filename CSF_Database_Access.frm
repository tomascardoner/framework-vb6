VERSION 5.00
Begin VB.Form CSF_Database_Access 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base de Datos"
   ClientHeight    =   2835
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
   Icon            =   "CSF_Database_Access.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBackupCopiesNumber 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3360
      TabIndex        =   8
      Tag             =   "INTEGER|EMPTY|NOTZERO|POSITIVE|99"
      Top             =   1620
      Width           =   390
   End
   Begin VB.CommandButton cmdDataSource 
      Caption         =   "..."
      Height          =   330
      Left            =   6960
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Examinar..."
      Top             =   180
      Width           =   255
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1140
      Width           =   3855
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   660
      Width           =   3855
   End
   Begin VB.TextBox txtDataSource 
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   180
      Width           =   3555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   6000
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4620
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
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
      TabIndex        =   7
      Top             =   1680
      Width           =   2730
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
      TabIndex        =   5
      Top             =   1200
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
      TabIndex        =   3
      Top             =   720
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
      TabIndex        =   0
      Top             =   240
      Width           =   1665
   End
End
Attribute VB_Name = "CSF_Database_Access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DUMMY_PASSWORD = "- ¿Creías que era tan fácil? -"

Private Const PROVIDER_MSACCESS = "Microsoft.Jet.OLEDB.4.0"
Private Const USER_DEFAULT_MSACCESS = "Admin"

Private mPassword As String

Private mKeyDecimal As Boolean

Private Sub Form_Load()
    Caption = App.Title & " - Base de Datos"
    
    Call CSM_Control_TextBox.PrepareAll(Me)
    
    txtDataSource.Text = pDatabase.DataSource
    txtUserID.Text = pDatabase.UserID
    
    txtUserID.Text = USER_DEFAULT_MSACCESS
    txtPassword.Text = ""
    
    mPassword = pDatabase.Password
    If mPassword <> "" Then
        txtPassword.Text = DUMMY_PASSWORD
    End If
    
    txtBackupCopiesNumber.Text = pDatabase.BackupCopiesNumber
    
    Call CSM_Control_TextBox.FormatAll(Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    mKeyDecimal = CSM_Control_TextBox.CheckKeyDown(ActiveControl, KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call CSM_Control_TextBox.CheckKeyPress(ActiveControl, KeyAscii, mKeyDecimal)
End Sub

Private Sub txtDataSource_GotFocus()
    CSM_Control_TextBox.SelAllText txtDataSource
End Sub

Private Sub cmdDataSource_Click()
    Dim Filename As String
    
    Filename = CSM_CommonDialog.FileOpen(Me.hWnd, "Seleccione la Base de Datos", "Bases de Datos de Microsoft Access (*.mdb)|*.mdb|Todos los Archivos (*.*)|*.*", IIf(txtDataSource.Text = "", App.Path, txtDataSource.Text))
    If Filename <> "" Then
        txtDataSource.Text = Filename
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

Private Sub cmdOK_Click()
    Dim DES As CSC_Encryption_DES
    
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
    
    pDatabase.Provider = PROVIDER_MSACCESS
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
    
    pDatabase.BackupCopiesNumber = Val(txtBackupCopiesNumber.Text)
    
    Call pDatabase.SaveParameters
    
    Me.Tag = "OK"
    Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CSF_Database_Access = Nothing
End Sub

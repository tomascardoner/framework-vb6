VERSION 5.00
Begin VB.Form CSF_DatabasePasswordChange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de la Contraseña de la Base de Datos"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4245
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDatabasePasswordConfirm 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   2595
   End
   Begin VB.TextBox txtDatabasePasswordNew 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2595
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1380
      TabIndex        =   6
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2760
      TabIndex        =   7
      Top             =   2220
      Width           =   1215
   End
   Begin VB.TextBox txtDatabasePasswordOld 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   2595
   End
   Begin VB.Label lblDatabasePasswordConfirm 
      AutoSize        =   -1  'True
      Caption         =   "Confirmación:"
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label lblDatabasePasswordNew 
      AutoSize        =   -1  'True
      Caption         =   "Nueva:"
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   510
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   120
      Picture         =   "CSF_DatabasePasswordChange.frx":0000
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lblDatabasePasswordOld 
      AutoSize        =   -1  'True
      Caption         =   "Anterior:"
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   630
   End
End
Attribute VB_Name = "CSF_DatabasePasswordChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim PasswordErrorCount As Long
    
    If txtDatabasePasswordOld.Text <> pDatabase.DatabasePassword Then
        PasswordErrorCount = PasswordErrorCount + 1
        MsgBox "La Contraseña Anterior es incorrecta.", vbExclamation, App.Title
        If PasswordErrorCount = 3 Then
            MsgBox "Ha realizado 3 intentos al ingresar la Contraseña Anterior. Se cerrará esta ventana.", vbCritical, App.Title
            Unload Me
            Exit Sub
        End If
        txtDatabasePasswordOld.SetFocus
        txtDatabasePasswordOld_GotFocus
        Exit Sub
    End If
    If txtDatabasePasswordNew.Text <> txtDatabasePasswordConfirm.Text Then
        MsgBox "La Nueva Contraseña y la Confirmación de la Contraseña no coinciden.", vbExclamation, App.Title
        txtDatabasePasswordConfirm.SetFocus
        txtDatabasePasswordConfirm_GotFocus
        Exit Sub
    End If
    pDatabase.DatabasePassword = txtDatabasePasswordNew.Text
    pDatabase.Connection.Close
    pDatabase.Connection.Properties("Jet OLEDB:New Database Password").Value = pDatabase.DatabasePassword
    pDatabase.Connection.Open
    
    Unload Me
End Sub

Private Sub txtDatabasePasswordOld_GotFocus()
    CSM_Control_TextBox.SelAllText txtDatabasePasswordOld
End Sub

Private Sub txtDatabasePasswordNew_GotFocus()
    CSM_Control_TextBox.SelAllText txtDatabasePasswordNew
End Sub

Private Sub txtDatabasePasswordConfirm_GotFocus()
    CSM_Control_TextBox.SelAllText txtDatabasePasswordConfirm
End Sub

VERSION 5.00
Begin VB.Form CSF_Database_Password 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingrese la Contraseña de la Base de Datos"
   ClientHeight    =   1545
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
   ScaleHeight     =   1545
   ScaleWidth      =   4245
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1380
      TabIndex        =   2
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2760
      TabIndex        =   3
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox txtDatabasePassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   2595
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   120
      Picture         =   "CSF_Database_Password.frx":0000
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblDatabasePassword 
      AutoSize        =   -1  'True
      Caption         =   "&Contraseña:"
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   885
   End
End
Attribute VB_Name = "CSF_Database_Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtDatabasePassword_GotFocus()
    CSM_Control_TextBox.SelAllText txtDatabasePassword
End Sub

Private Sub cmdOK_Click()
    pDatabase.DatabasePassword = txtDatabasePassword.Text
    Tag = "OK"
    Hide
End Sub

Private Sub cmdCancel_Click()
    Tag = "CANCEL"
    Hide
End Sub

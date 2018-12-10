VERSION 5.00
Begin VB.Form CSF_Database_ReplicationPath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bases de datos replicadas"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CSF_Database_ReplicationPath.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   8460
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkEjectDrive 
      Caption         =   "Expulsar la unidad al finalizar."
      Height          =   210
      Left            =   1980
      TabIndex        =   5
      Top             =   660
      Value           =   1  'Checked
      Width           =   2715
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5700
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   7080
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemoteDB 
      Caption         =   "..."
      Height          =   330
      Left            =   8040
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Examinar..."
      Top             =   180
      Width           =   255
   End
   Begin VB.ComboBox cboRemoteDB 
      Height          =   330
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   6015
   End
   Begin VB.Label lblRemoteDB 
      AutoSize        =   -1  'True
      Caption         =   "Base de datos remota:"
      Height          =   210
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1635
   End
End
Attribute VB_Name = "CSF_Database_ReplicationPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Saved_ListIndex As Long
    
    Call CSM_Control_ComboBox.FillFromSQL(cboRemoteDB, "SELECT DISTINCT Pathname FROM MSysReplicas WHERE Pathname IS NOT NULL ORDER BY Pathname", "", "Pathname", "Bases de datos replicadas")
    Saved_ListIndex = CSM_Control_ComboBox.GetListIndexByText(cboRemoteDB, CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "Sync_RemoteDBPath", "", csrdtString), cscpItemOrNone)
    If Saved_ListIndex > -1 Then
        cboRemoteDB.ListIndex = Saved_ListIndex
    End If
    chkEjectDrive.Value = IIf(CSM_Registry.GetValue_FromApplication_LocalMachine("Database", "Sync_EjectUSBDrive", True, csrdtBoolean), vbChecked, vbUnchecked)
End Sub

Private Sub cmdRemoteDB_Click()
    Dim Filename As String
    
    Filename = CSM_CommonDialog.FileOpen(Me.hwnd, "Seleccione la Base de Datos", "Bases de Datos de Microsoft Access (*.mdb)|*.mdb|Todos los Archivos (*.*)|*.*", IIf(cboRemoteDB.Text = "", App.Path, cboRemoteDB.Text))
    If Filename <> "" Then
        cboRemoteDB.AddItem Filename
        cboRemoteDB.ListIndex = cboRemoteDB.NewIndex
    End If
End Sub

Private Sub cmdOK_Click()
    If cboRemoteDB.ListIndex = -1 Then
        MsgBox "Debe seleccionar la base de datos remota.", vbInformation, App.Title
        cboRemoteDB.SetFocus
        Exit Sub
    End If
    
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "Sync_RemoteDBPath", cboRemoteDB.Text)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Database", "Sync_EjectUSBDrive", (chkEjectDrive.Value = vbChecked))
    
    Me.Tag = "OK"
    Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub

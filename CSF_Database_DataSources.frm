VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form CSF_Database_DataSources 
   Caption         =   "Seleccione el Origen de los datos"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CSF_Database_DataSources.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwData 
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Nombre"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Location"
         Text            =   "Ubicación"
         Object.Width           =   9525
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   5880
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "CSF_Database_DataSources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Index As Integer
    Dim ListItem As ListItem
    
    For Index = 1 To pDatabase.CDataSourcesNames.Count
        Set ListItem = lvwData.ListItems.Add(, pDatabase.CDataSources(Index), pDatabase.CDataSourcesNames(Index))
        ListItem.SubItems(1) = pDatabase.CDataSources(Index)
    Next Index
End Sub

Private Sub lvwData_DblClick()
    Call cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    If lvwData.SelectedItem Is Nothing Then
        MsgBox "Debe seleccionar el Origen de los datos.", vbInformation, App.Title
        lvwData.SetFocus
        Exit Sub
    End If
    
    pDatabase.DataSource = lvwData.SelectedItem.Key
    
    Me.Tag = "OK"
    Hide
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub


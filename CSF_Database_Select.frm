VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form CSF_Database_Select 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccione la Base de Datos"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CSF_Database_Select.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwData 
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5644
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2580
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "CSF_Database_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowAndSelect(ByVal FileNameFilter As String) As String
    Dim FileName As String
    
    FileName = Dir(CSM_String.RemoveLastSubString(pDatabase.DataSource, "\") & "\" & FileNameFilter, vbArchive)
    
    If FileName <> "" Then
        Do While FileName <> ""
            lvwData.ListItems.Add , FileName, CSM_String.RemoveLastSubString(FileName, ".")
            FileName = Dir
        Loop
        If lvwData.ListItems.Count = 1 Then
            ShowAndSelect = lvwData.ListItems(1).Key
        Else
            Screen.MousePointer = vbDefault
            DoEvents
            CSF_Database_Select.Show vbModal
            If CSF_Database_Select.Tag = "OK" Then
                ShowAndSelect = lvwData.SelectedItem.Key
            End If
        End If
    End If
End Function

Private Sub cmdOK_Click()
    If lvwData.SelectedItem Is Nothing Then
        MsgBox "Debe seleccionar una Base de Datos.", vbInformation, App.Title
        lvwData.SetFocus
        Exit Sub
    End If
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

Private Sub lvwData_DblClick()
    cmdOK_Click
End Sub

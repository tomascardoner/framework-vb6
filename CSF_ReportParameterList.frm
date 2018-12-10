VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form CSF_ReportParameterList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros del Reporte"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5820
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7140
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   4395
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   7752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Descripcion"
         Text            =   "Descripción"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Tipo"
         Text            =   "Tipo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Valor"
         Text            =   "Valor"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "CSF_ReportParameterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mReportParameters As Collection


Public Function LoadListData(ByVal ReportParameters As Collection) As Boolean
    Dim ReportParameter As CSC_ReportParameter
    Dim ListItem As MSComctlLib.ListItem
    
    If pIsCompiled Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set mReportParameters = ReportParameters
    
    lvwData.ListItems.Clear
    
    For Each ReportParameter In ReportParameters
        If ReportParameter.ParameterAskFor Then
            Set ListItem = lvwData.ListItems.Add(, ReportParameter.ParameterName, ReportParameter.ParameterDescription)
            ListItem.SubItems(1) = IIf(ReportParameter.ParameterRequired, "Requerido", "Opcional")
            Call ShowItemValueData(ListItem, ReportParameter)
        End If
    Next ReportParameter
    
    Screen.MousePointer = vbDefault
    LoadListData = True
    Exit Function
    
ErrorHandler:
    CSM_Error.ShowErrorMessage "Forms.ReportParameterList.LoadListData", "Error al cargar la lista de Parámetros del Reporte."
End Function

Private Sub ShowItemValueData(ByRef ListItem As MSComctlLib.ListItem, ByVal ReportParameter As CSC_ReportParameter)
    With ReportParameter
        If IsEmpty(.ParameterValue) Then
            ListItem.SubItems(2) = " "
        Else
            Select Case .ParameterDataType
                Case csrpdtUndefined
                    ListItem.SubItems(2) = "« Undefined data type »"
                Case csrpdtNumberInteger
                    ListItem.SubItems(2) = Format(CLng(.ParameterValue), "#.###")
                Case csrpdtNumberDecimal
                    ListItem.SubItems(2) = Format(CDec(.ParameterValue), "#.###,##")
                Case csrpdtCurrency
                    ListItem.SubItems(2) = Format(CCur(.ParameterValue), "Currency")
                Case csrpdtDate
                    ListItem.SubItems(2) = Format(CDate(.ParameterValue), "Short Date")
                Case csrpdtTime
                    ListItem.SubItems(2) = Format(CDate(.ParameterValue), "Short Time")
                Case csrpdtDateTime
                    ListItem.SubItems(2) = Format(CDate(.ParameterValue), "Short Date") & " " & Format(.ParameterValue, "Short Time")
                Case csrpdtString
                    ListItem.SubItems(2) = CStr(.ParameterValue)
                Case csrpdtBoolean
                    ListItem.SubItems(2) = IIf(CBool(.ParameterValue), "Sí", "No")
                Case csrpdtList, csrpdtWeekday, csrpdtMonth
                    ListItem.SubItems(2) = .ParameterDisplayValue
                Case csrpdtYear
                    ListItem.SubItems(2) = .ParameterValue
            End Select
        End If
    End With
End Sub

Private Sub Form_Load()
    Call ResizeAndPosition(frmMDI, Me)
End Sub

Private Sub lvwData_DblClick()
    Load CSF_ReportParameterDetail
    CSF_ReportParameterDetail.SetData mReportParameters(lvwData.SelectedItem.Key)
    CSF_ReportParameterDetail.Show vbModal, CSF_ReportParameterList
    If CSF_ReportParameterDetail.Tag = "OK" Then
        Call ShowItemValueData(lvwData.SelectedItem, mReportParameters(lvwData.SelectedItem.Key))
    End If
    Unload CSF_ReportParameterDetail
    Set CSF_ReportParameterDetail = Nothing
End Sub

Private Sub cmdAceptar_Click()
    Dim ReportParameter As CSC_ReportParameter
    
    For Each ReportParameter In mReportParameters
        If ReportParameter.ParameterAskFor And ReportParameter.ParameterRequired Then
            If IsEmpty(ReportParameter.ParameterValue) Then
                MsgBox "Debe especificar un valor para el Parámetro requerido." & vbCr & vbCr & ReportParameter.ParameterDescription, vbExclamation, App.Title
                Set lvwData.SelectedItem = lvwData.ListItems(ReportParameter.ParameterName)
                Call lvwData.ListItems(ReportParameter.ParameterName).EnsureVisible
                lvwData.SetFocus
                Exit Sub
            End If
        End If
    Next ReportParameter
    
    Tag = "OK"
    Hide
End Sub

Private Sub cmdCancelar_Click()
    Tag = "CANCEL"
    Hide
End Sub

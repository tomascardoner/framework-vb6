VERSION 5.00
Begin VB.Form CSF_ReportParameterDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle del Parámetro del Reporte"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
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
   ScaleHeight     =   2085
   ScaleWidth      =   5250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNotNull 
      Height          =   330
      Left            =   180
      TabIndex        =   4
      Top             =   720
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.ComboBox cboValue 
      Height          =   330
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3900
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2580
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblDescripition 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4875
   End
End
Attribute VB_Name = "CSF_ReportParameterDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mReportParameter As CSC_ReportParameter

Public Function SetData(ByRef ReportParameter As CSC_ReportParameter) As Boolean
    Dim ListIndex As Integer
    
    Set mReportParameter = ReportParameter

    lblDescripition.Caption = mReportParameter.ParameterDescription
    Select Case mReportParameter.ParameterDataType
        Case csrpdtUndefined
        Case csrpdtString
        Case csrpdtNumberInteger
        Case csrpdtNumberDecimal
        Case csrpdtCurrency
        Case csrpdtDate
        Case csrpdtTime
        Case csrpdtDateTime
        Case csrpdtBoolean
            cboValue.Clear
            cboValue.AddItem "Sí"
            cboValue.AddItem "No"
            If IsEmpty(mReportParameter.ParameterValue) Then
                cboValue.ListIndex = -1
            Else
                cboValue.ListIndex = 1 + CInt(mReportParameter.ParameterValue)
            End If
        Case csrpdtList
            Call CSM_Control_ComboBox.FillFromSQL(cboValue, mReportParameter.ParameterListValuesOrRecordSource, mReportParameter.ParameterListFieldNameBound, mReportParameter.ParameterListFieldNameDisplay, mReportParameter.ParameterListErrorEntityName, cscpItemOrFirstIfUnique, mReportParameter.ParameterValue)
        Case csrpdtWeekday
            cboValue.Clear
            For ListIndex = 1 To 7
                cboValue.AddItem WeekdayName(ListIndex)
            Next ListIndex
            cboValue.ListIndex = mReportParameter.ParameterValue - 1
        Case csrpdtMonth
            cboValue.Clear
            For ListIndex = 1 To 12
                cboValue.AddItem MonthName(ListIndex)
            Next ListIndex
            cboValue.ListIndex = mReportParameter.ParameterValue - 1
        Case csrpdtYear
            If IsEmpty(mReportParameter.ParameterMinValue) Then
                mReportParameter.ParameterMinValue = Year(Date)
            End If
            If IsEmpty(mReportParameter.ParameterMaxValue) Then
                mReportParameter.ParameterMaxValue = Year(Date) + 20
            End If
            cboValue.Clear
            For ListIndex = mReportParameter.ParameterMinValue To mReportParameter.ParameterMaxValue
                cboValue.AddItem ListIndex
                cboValue.ItemData(cboValue.NewIndex) = ListIndex
            Next ListIndex
            cboValue.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(cboValue, mReportParameter.ParameterValue)
    End Select
    
    SetData = True
End Function

Private Sub chkNotNull_Click()
    cboValue.Enabled = (chkNotNull.Value = vbChecked)
End Sub

Private Sub cmdAceptar_Click()
    If chkNotNull.Value = vbChecked Then
        Select Case mReportParameter.ParameterDataType
            Case csrpdtUndefined
            Case csrpdtString
            Case csrpdtNumberInteger
            Case csrpdtNumberDecimal
            Case csrpdtCurrency
            Case csrpdtDate
            Case csrpdtTime
            Case csrpdtDateTime
            Case csrpdtBoolean
                If cboValue.ListIndex = -1 Then
                    MsgBox "Debe seleccionar un item de la Lista.", vbExclamation, App.Title
                    cboValue.SetFocus
                    Exit Sub
                End If
                mReportParameter.ParameterValue = CBool(-1 + cboValue.ListIndex)
            Case csrpdtList, csrpdtYear
                If cboValue.ListIndex = -1 Then
                    MsgBox "Debe seleccionar un item de la Lista.", vbExclamation, App.Title
                    cboValue.SetFocus
                    Exit Sub
                End If
                mReportParameter.ParameterValue = cboValue.ItemData(cboValue.ListIndex)
                mReportParameter.ParameterDisplayValue = cboValue.Text
            Case csrpdtWeekday, csrpdtMonth, csrpdtYear
                If cboValue.ListIndex = -1 Then
                    MsgBox "Debe seleccionar un item de la Lista.", vbExclamation, App.Title
                    cboValue.SetFocus
                    Exit Sub
                End If
                mReportParameter.ParameterValue = cboValue.ListIndex + 1
                mReportParameter.ParameterDisplayValue = cboValue.Text
        End Select
    Else
        mReportParameter.ParameterValue = Null
    End If
    
    Tag = "OK"
    Hide
End Sub

Private Sub cmdCancelar_Click()
    Tag = "CANCEL"
    Hide
End Sub

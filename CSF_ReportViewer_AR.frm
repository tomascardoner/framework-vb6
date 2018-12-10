VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form CSF_Report_AR_Viewer 
   Caption         =   "Print Report"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   Icon            =   "CSF_ReportViewer_AR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewer 
      Height          =   4635
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8176
      SectionData     =   "CSF_ReportViewer_AR.frx":062A
   End
End
Attribute VB_Name = "CSF_Report_AR_Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call CSM_Forms.ResizeAndPosition(frmMDI, Me)
End Sub

Private Sub Form_Resize()
    Const CONTROL_SPACE = 60
    
    With ARViewer
        .Top = CONTROL_SPACE
        .Left = CONTROL_SPACE
        .Height = ScaleHeight - (CONTROL_SPACE * 2)
        .Width = ScaleWidth - (CONTROL_SPACE * 2)
    End With
End Sub

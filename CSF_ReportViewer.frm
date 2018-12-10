VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form CSF_ReportViewer 
   Caption         =   "Reportes"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   Icon            =   "CSF_ReportViewer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   9465
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   4935
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9315
      _cx             =   16431
      _cy             =   8705
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   11274
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "CSF_ReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CRAXDRTReport As CRAXDRT.Report

Private Sub Form_Load()
    CSM_Forms.ResizeAndPosition frmMDI, Me
End Sub

Private Sub Form_Resize()
    Const CONTROL_SPACE = 60
    
    On Error Resume Next
    
    CRViewer.Top = CONTROL_SPACE
    CRViewer.Left = CONTROL_SPACE
    CRViewer.Height = ScaleHeight - (CONTROL_SPACE * 2)
    CRViewer.Width = ScaleWidth - (CONTROL_SPACE * 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CRAXDRTReport = Nothing
End Sub

Public Sub PrinterSetup()
    CRAXDRTReport.PrinterSetup frmMDI.hwnd
End Sub

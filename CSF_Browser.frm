VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form CSF_Browser 
   Caption         =   "Navegador"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CSF_Browser.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser webMain 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   7095
      ExtentX         =   12515
      ExtentY         =   5953
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin ComCtl3.CoolBar cbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   741
      BandCount       =   1
      FixedOrder      =   -1  'True
      BandBorders     =   0   'False
      _CBWidth        =   7350
      _CBHeight       =   420
      _Version        =   "6.7.9782"
      MinHeight1      =   360
      Width1          =   2880
      FixedBackground1=   0   'False
      UseCoolbarColors1=   0   'False
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4035
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CSF_Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Paint()
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    Const MARGIN As Long = 120
    
    DoEvents
    
    webMain.Top = cbrMain.Height + MARGIN
    webMain.Left = MARGIN
    webMain.Height = stbMain.Top - MARGIN - webMain.Top
    webMain.Width = Me.ScaleWidth - MARGIN - webMain.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set CSF_Browser = Nothing
End Sub

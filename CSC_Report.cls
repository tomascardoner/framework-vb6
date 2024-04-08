VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFileName As String
Private mWindowTitle As String
Private mDisplayGroupTree As Boolean
Private mRecordSelectionFormula As String
Private mParentForm_hWnd As Long

'PRINTER PROPERTIES
Private mPrinterSetup As Boolean
Private mPrinterName As String
Private mPrinterBin As Integer
Private mPrinterPaper As Integer
Private mPaperSizeLengthInPixels As Integer
Private mPaperSizeWidthInPixels As Integer

Private mCParameters As Collection
Private mRecordset As ADODB.recordset

Private mCRAXDRTReport As CRAXDRT.Report

Public Property Get FILENAME() As String
    FILENAME = mFileName
End Property

Public Property Let FILENAME(ByVal value As String)
    mFileName = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get WindowTitle() As String
    WindowTitle = mWindowTitle
End Property

Public Property Let WindowTitle(ByVal value As String)
    mWindowTitle = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ParentForm_hWnd() As Long
    ParentForm_hWnd = mParentForm_hWnd
End Property

Public Property Let ParentForm_hWnd(ByVal value As Long)
    mParentForm_hWnd = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PaperSizeLengthInPixels() As Integer
    PaperSizeLengthInPixels = mPaperSizeLengthInPixels
End Property

Public Property Let PaperSizeLengthInPixels(ByVal value As Integer)
    mPaperSizeLengthInPixels = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get PaperSizeWidthInPixels() As Integer
    PaperSizeWidthInPixels = mPaperSizeWidthInPixels
End Property

Public Property Let PaperSizeWidthInPixels(ByVal value As Integer)
    mPaperSizeWidthInPixels = value
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Parameter(ByVal ParameterName As String) As CSC_ReportParameter
    Set Parameter = mCParameters(ParameterName)
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function ParameterAdd(ByVal ParameterName As String, ByVal ParameterDescription As String, ByVal ParameterDataType As csReportParameterDataType, ByVal ParameterAllowMoreThanOne As Boolean, ByVal ParameterAllowRange As Boolean, ByVal ParameterRequired As Boolean, Optional ByVal ParameterListValuesOrRecordSource As String = "", Optional ByVal ParameterListFieldNameBound As String = "", Optional ByVal ParameterListFieldNameDisplay As String = "", Optional ByVal ParameterListErrorEntityName As String = "", Optional ByVal ParameterMinValue As Variant = Empty, Optional ByVal ParameterMaxValue As Variant = Empty, Optional ByVal ParameterValue As Variant = Empty, Optional ByVal ParameterDisplayValue As String = "", Optional ParameterAskFor As Boolean = True) As CSC_ReportParameter
    Dim NewParameter As CSC_ReportParameter
    
    Set NewParameter = New CSC_ReportParameter
    With NewParameter
        .ParameterName = ParameterName
        .ParameterDescription = ParameterDescription
        .ParameterDataType = ParameterDataType
        .ParameterAllowMoreThanOne = ParameterAllowMoreThanOne
        .ParameterAllowRange = ParameterAllowRange
        .ParameterRequired = ParameterRequired
        .ParameterListValuesOrRecordSource = ParameterListValuesOrRecordSource
        .ParameterListFieldNameBound = ParameterListFieldNameBound
        .ParameterListFieldNameDisplay = ParameterListFieldNameDisplay
        .ParameterListErrorEntityName = ParameterListErrorEntityName
        .ParameterMinValue = ParameterMinValue
        .ParameterMaxValue = ParameterMaxValue
        .ParameterDisplayValue = ParameterDisplayValue
        .ParameterValue = ParameterValue
        .ParameterAskFor = ParameterAskFor
    End With
    mCParameters.Add NewParameter, ParameterName
    Set ParameterAdd = NewParameter
    Set NewParameter = Nothing
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Parameters() As Collection
    Set Parameters = mCParameters
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function AskForParametersAsDialog(ByRef ParentForm As Form) As Boolean
    Dim ReportParameter As CSC_ReportParameter
    Dim ShowForm As Boolean
    
    If mCParameters.Count > 0 Then
        For Each ReportParameter In mCParameters
            If ReportParameter.ParameterAskFor Then
                ShowForm = True
                Exit For
            End If
        Next ReportParameter
        
        If ShowForm Then
            Load CSF_ReportParameterList
            
            If CSF_ReportParameterList.LoadListData(mCParameters) Then
                Call CSF_ReportParameterList.Show(vbModal, ParentForm)
                AskForParametersAsDialog = (CSF_ReportParameterList.Tag = "OK")
            End If
            
            Unload CSF_ReportParameterList
            Set CSF_ReportParameterList = Nothing
        Else
            AskForParametersAsDialog = True
        End If
    Else
        AskForParametersAsDialog = True
    End If
End Function

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get ExportOptions() As CRAXDRT.ExportOptions
    Set ExportOptions = mCRAXDRTReport.ExportOptions
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get Report() As CRAXDRT.Report
    Set Report = mCRAXDRTReport
End Property

Public Property Let Report(ByRef rpt As CRAXDRT.Report)
    Set mCRAXDRTReport = rpt
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Get recordset() As ADODB.recordset
    Set recordset = mRecordset
End Property

Public Property Set recordset(ByVal value As ADODB.recordset)
    Set mRecordset = value.Clone
End Property

'////////////////////////////////////////////////////////////////////////
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function OpenReport(ByVal IsSQLServer As Boolean) As Boolean
    Dim errorMessage As String
    
    Dim ParameterCurrent As CSC_ReportParameter
    Dim ParameterFieldDefinition As CRAXDRT.ParameterFieldDefinition

    Dim DatabaseTable As CRAXDRT.DatabaseTable
    Dim Section As CRAXDRT.Section
    
    Dim ReportObject As Object
    Dim SubReportObject As CRAXDRT.SubReportObject
    
    Dim PrinterCurrent As Printer
    Dim strDriverName As String
    Dim strDeviceName As String
    Dim strPort As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Start the Application if its the first time
    If pCRAXDRTApplication Is Nothing Then
        errorMessage = "Ha ocurrido un error al inciar una nueva sesión de Crystal Reports."
        Set pCRAXDRTApplication = New CRAXDRT.Application
    End If
    
    errorMessage = "Error al abrir el Reporte." & vbCr & vbCr & "FileName: " & mFileName
    Set mCRAXDRTReport = pCRAXDRTApplication.OpenReport(mFileName, crOpenReportByDefault)
    
    With mCRAXDRTReport
        '//////////////////////////////////////////////////////
        'CHANGE PRINTER IF SPECIFIED
        If mPrinterName <> "" Then
            'Find specified details so driver and port can be set correctly
            For Each PrinterCurrent In Printers
                If PrinterCurrent.DeviceName = mPrinterName Then
                    Set Printer = PrinterCurrent
                    strDriverName = PrinterCurrent.DriverName
                    strDeviceName = PrinterCurrent.DeviceName
                    strPort = PrinterCurrent.Port
                    
                    .SelectPrinter strDriverName, strDeviceName, strPort
                    Exit For
                End If
            Next PrinterCurrent
        End If
        
        '.PaperSource = varBin
    
        '//////////////////////////////////////////////////////
        'SET USER DEFINED SIZE
        If mPaperSizeLengthInPixels <> 0 And mPaperSizeWidthInPixels <> 0 Then
            mCRAXDRTReport.PaperSize = crPaperUser
            mCRAXDRTReport.SetUserPaperSize 2, 2  'mPaperSizeLengthInPixels, mPaperSizeWidthInPixels
            mCRAXDRTReport.PaperSize = crPaperUser
        End If
        
        '//////////////////////////////////////////////////////
        'SET DATABASE INFO
        If Not recordset Is Nothing Then
            errorMessage = "Error al asignar el recordset al Reporte."
            Call .database.SetDataSource(mRecordset)
        Else
            For Each DatabaseTable In Report.database.Tables
                errorMessage = "Error al asignar la base de datos a la Tabla '" & DatabaseTable.Name & "' del Reporte."
                DatabaseTable.ConnectionProperties("Provider").value = pDatabase.Provider
                DatabaseTable.ConnectionProperties("Data Source").value = pDatabase.DataSource
                If IsSQLServer Then
                    DatabaseTable.ConnectionProperties("Initial Catalog").value = pDatabase.database
                    DatabaseTable.ConnectionProperties("Connect Timeout").value = pDatabase.ConnectionTimeout
                    DatabaseTable.ConnectionProperties("General Timeout").value = pDatabase.CommandTimeout
                End If
                DatabaseTable.ConnectionProperties("Password").value = pDatabase.Password
                If IsSQLServer Then
                    DatabaseTable.Location = pDatabase.database & ".dbo." & DatabaseTable.Location
                End If
            Next DatabaseTable
        
            For Each Section In Report.Sections
                For Each ReportObject In Section.ReportObjects
                    If ReportObject.Kind = crSubreportObject Then
                        Set SubReportObject = ReportObject
                        For Each DatabaseTable In SubReportObject.OpenSubreport.database.Tables
                            errorMessage = "Error al asignar la base de datos a la Tabla '" & DatabaseTable.Name & "' del Sub-Reporte '" & SubReportObject.Name & "'."
                            'DatabaseTable.ConnectionProperties("Provider").Value = pParametro.Report_Provider
                            DatabaseTable.ConnectionProperties("Data Source").value = pDatabase.DataSource
                            If IsSQLServer Then
                                DatabaseTable.ConnectionProperties("Initial Catalog").value = pDatabase.database
                                DatabaseTable.ConnectionProperties("Connect Timeout").value = pDatabase.ConnectionTimeout
                                DatabaseTable.ConnectionProperties("General Timeout").value = pDatabase.CommandTimeout
                            End If
                            DatabaseTable.ConnectionProperties("Password").value = pDatabase.Password
                            If IsSQLServer Then
                                DatabaseTable.Location = pDatabase.database & ".dbo." & DatabaseTable.Location
                            End If
                        Next DatabaseTable
                    End If
                Next ReportObject
            Next Section
        End If
        
        '//////////////////////////////////////////////////////
        'SET PARAMETERS VALUES
        For Each ParameterCurrent In mCParameters
            If IsNull(ParameterCurrent.ParameterValue) Or IsEmpty(ParameterCurrent.ParameterValue) Then
            Else
                errorMessage = "Error al definir el Parámetro del Reporte." & vbCr & vbCr & "Parametro: " & ParameterCurrent.ParameterName
                Set ParameterFieldDefinition = mCRAXDRTReport.ParameterFields.GetItemByName(IIf(Left(ParameterCurrent.ParameterName, 1) = "@", "", "@") & ParameterCurrent.ParameterName)
                With ParameterFieldDefinition
                    Select Case .ValueType
                        Case crInt8sField, crInt8uField, crInt16sField, crInt16uField, crInt32sField, crInt32uField, crNumberField
                            .AddCurrentValue CLng(ParameterCurrent.ParameterValue)
                        Case crCurrencyField
                            .AddCurrentValue CDbl(ParameterCurrent.ParameterValue)
                        Case crBooleanField
                            .AddCurrentValue CBool(ParameterCurrent.ParameterValue)
                        Case crStringField, crTransientMemoField, crPersistentMemoField, crBlobField
                            .AddCurrentValue CStr(ParameterCurrent.ParameterValue)
                        Case crDateField, crTimeField, crDateTimeField
                            .AddCurrentValue CDate(ParameterCurrent.ParameterValue)
                        Case crBitmapField, crIconField, crPictureField, crOleField, crChartField
                        Case crUnknownField
                    End Select
                End With
                Set ParameterFieldDefinition = Nothing
            End If
        Next ParameterCurrent
        
        
        '//////////////////////////////////////////////////////
        'SET APPLICATION INFO
        .ApplicationName = App.Title
        .SetDialogParentWindow ParentForm_hWnd
        If mRecordSelectionFormula <> "" Then
            .recordSelectionFormula = mRecordSelectionFormula
        End If
        If mPrinterSetup Then
            .PrinterSetup ParentForm_hWnd
        End If
        
        .EnableParameterPrompting = False
    End With
    Screen.MousePointer = vbDefault
    OpenReport = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CSC_Report.OpenReport", errorMessage
End Function

Public Function PreviewReport(ByVal promptUser As Boolean) As Boolean
    Dim ReportViewer As CSF_ReportViewer
    Dim FormLoaded As Boolean
    Dim errorMessage As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    errorMessage = "Error al abrir una nueva sesión del Visor de Reportes."
    Set ReportViewer = New CSF_ReportViewer
    Load ReportViewer
    FormLoaded = True
    With ReportViewer.CRViewer
        .DisplayGroupTree = mDisplayGroupTree
        errorMessage = "Error al abrir el Reporte."
        .ReportSource = mCRAXDRTReport
        If promptUser Then
            Call mCRAXDRTReport.PrinterSetupEx(mParentForm_hWnd)
        End If
        .ViewReport
    End With
    ReportViewer.Caption = "REPORTE: " & mWindowTitle
    ReportViewer.Show
    
    Screen.MousePointer = vbDefault

    PreviewReport = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CSC_Report.PreviewReport", errorMessage
    If FormLoaded Then
        Unload ReportViewer
        Set ReportViewer = Nothing
    End If
End Function

Public Function PrintReport(ByVal promptUser As Boolean, Optional ByVal NumberOfCopies, Optional ByVal collatedCopies, Optional ByVal StartPage, Optional ByVal stopPage) As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Call mCRAXDRTReport.PrintOut(promptUser, NumberOfCopies, collatedCopies, StartPage, stopPage)
    Screen.MousePointer = vbDefault
    
    PrintReport = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CSC_Report.PrintReport", "Error al imprimir el Reporte."
End Function

Public Function ExportReport(Optional ByVal promptUser As Boolean = False) As Boolean
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    mCRAXDRTReport.Export promptUser

    Screen.MousePointer = vbDefault
    ExportReport = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Classes.CSC_Report.Export", "Error al Exportar el Reporte."
End Function

Private Sub Class_Initialize()
    Set mCParameters = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCParameters = Nothing
    If Not mRecordset Is Nothing Then
        If mRecordset.State = adStateOpen Then
            If Not (mRecordset.BOF Or mRecordset.EOF) Then
                If mRecordset.EditMode <> adEditNone Then
                    mRecordset.CancelUpdate
                End If
            End If
            mRecordset.Close
        End If
        Set mRecordset = Nothing
    End If
End Sub

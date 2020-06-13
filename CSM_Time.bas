Attribute VB_Name = "CSM_Time"
Option Explicit

'=========================================================================
'MULTIMEDIA TIMER
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'=========================================================================

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_TIMECHANGE = &H1E
Private Const HWND_TOPMOST = -1

Public Sub SyncronizeWithSQLServer(Optional ByVal SyncEvenIfLocal As Boolean = False)
    Dim recData As ADODB.Recordset
    
    On Error Resume Next
    
    If Not SyncEvenIfLocal Then
        If CSM_Session.GetComputerName() = pDatabase.DataSource Then
            WriteLogEvent "Synchronizing Date/Time with SQL Server - SKIPPED because this machine is the server.", vbLogEventTypeInformation, pParametro.LogAccion_Enabled
            Exit Sub
        End If
    End If
    
    Set recData = New ADODB.Recordset
    recData.Source = "SELECT convert(char(10), getdate(), 111) AS Fecha, convert(char(8), getdate(), 108) AS Hora"
    recData.Open , pDatabase.Connection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    WriteLogEvent "Synchronizing Date/Time with SQL Server - OLD: " & Format(Date, "Short Date") & " " & Format(Time, "Short Time") & " - NEW: " & Format(CDate(recData("Fecha").Value), "Short Date") & " " & Format(CDate(recData("Hora").Value), "Short Time"), vbLogEventTypeInformation, pParametro.LogAccion_Enabled
    
    Date = CDate(recData("Fecha").Value)
    Time = CDate(recData("Hora").Value)
    
    recData.Close
    Set recData = Nothing
    
    Call SendMessage(HWND_TOPMOST, WM_TIMECHANGE, 0, 0)
End Sub

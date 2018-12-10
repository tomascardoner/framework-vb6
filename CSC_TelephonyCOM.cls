VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_TelephonyCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLocationPulse As Boolean
Private mLocationCountryCode As Integer
Private mLocationAreaCode As String
Private mLocationLocalAccessCode As String
Private mLocationLongDistanceAccessCode As String

Private mCSerialCOMIDs As Collection
Private mCModemCOMPorts As Collection
Private mCModemNames As Collection

Private mCallerIDNumber As String
Private mCallerIDDateTime As Date

Public Event Ringing()
Public Event Identifying()

Public Property Get LocationPulse() As Boolean
    LocationPulse = mLocationPulse
End Property

Public Property Get LocationCountryCode() As Integer
    LocationCountryCode = mLocationCountryCode
End Property

Public Property Get LocationAreaCode() As String
    LocationAreaCode = mLocationAreaCode
End Property

Public Property Get LocationLocalAccessCode() As String
    LocationLocalAccessCode = mLocationLocalAccessCode
End Property

Public Property Get LocationLongDistanceAccessCode() As String
    LocationLongDistanceAccessCode = mLocationLongDistanceAccessCode
End Property

Public Function LoadLocationInfo() As Boolean
    Dim LocationDefault As Byte
    
    'LOCATION INFO
    LocationDefault = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Locations", "CurrentID", 1, csrdtNumberInteger)
    'mLocationPulse =
    mLocationCountryCode = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Locations\Location" & LocationDefault, "Country", 54, csrdtNumberInteger)
    mLocationAreaCode = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Locations\Location" & LocationDefault, "AreaCode", "02227", csrdtString)
    mLocationLocalAccessCode = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Locations\Location" & LocationDefault, "OutsideAccess", "", csrdtString)
    mLocationLongDistanceAccessCode = CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Telephony\Locations\Location" & LocationDefault, "LongDistanceAccess", "", csrdtString)
    
    LoadLocationInfo = True
End Function

Public Property Get CSerialCOMIDs() As Collection
    Set CSerialCOMIDs = mCSerialCOMIDs
End Property

Public Property Get CModemCOMPorts() As Collection
    Set CModemCOMPorts = mCModemCOMPorts
End Property

Public Property Get CModemNames() As Collection
    Set CModemNames = mCModemNames
End Property

Public Property Get CallerIDNumber() As String
    CallerIDNumber = mCallerIDNumber
End Property

Public Property Get CallerIDDateTime() As Date
    CallerIDDateTime = mCallerIDDateTime
End Property

Public Function LoadPortsList() As Boolean
    Dim CValueTypes As Collection
    Dim ModemCOMPortIndex As Integer
    Dim ModemCOMPort As String
    Dim CModemsIndex As Collection
    Dim ModemIndex As Variant
    Dim COMPortFound As Boolean
    
    Const SERIAL_PORTS_ENUM_KEY As String = "HARDWARE\DEVICEMAP\SERIALCOMM"
    Const MODEMS_ENUM_KEY As String = "SYSTEM\CurrentControlSet\Control\Class\{4D36E96D-E325-11CE-BFC1-08002BE10318}"
    
    'OBTENGO LA LISTA DE PUERTOS COMM HABILITADOS
    Set mCSerialCOMIDs = New Collection
    Call CSM_Registry.EnumerateValues(csrkHKEY_LOCAL_MACHINE, SERIAL_PORTS_ENUM_KEY, mCSerialCOMIDs, CValueTypes, mCModemCOMPorts, "", False)
    
    'BUSCO LOS MODEMS CORRESPONDIENTES
    Set mCModemNames = New Collection
    Set CModemsIndex = CSM_Registry.EnumerateSubKeys(csrkHKEY_LOCAL_MACHINE, MODEMS_ENUM_KEY)
    
    ModemCOMPortIndex = 1
    Do While ModemCOMPortIndex <= mCModemCOMPorts.Count
        COMPortFound = False
        ModemCOMPort = mCModemCOMPorts(ModemCOMPortIndex)
        For Each ModemIndex In CModemsIndex
            If CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, MODEMS_ENUM_KEY & "\" & ModemIndex, "AttachedTo", "", csrdtString) = ModemCOMPort Then
                mCModemNames.Add CSM_Registry.GetValue(csrkHKEY_LOCAL_MACHINE, MODEMS_ENUM_KEY & "\" & ModemIndex, "FriendlyName", "", csrdtString)
                COMPortFound = True
                Exit For
            End If
        Next ModemIndex
        If COMPortFound Then
            ModemCOMPortIndex = ModemCOMPortIndex + 1
        Else
            mCModemCOMPorts.Remove ModemCOMPortIndex
        End If
    Loop
    
    LoadPortsList = True
End Function

Public Function CallerIDCheck(ByRef comControl As MSCommLib.MSComm) As Boolean
    Dim BufferCurrent As String
    Dim Buffer As String
    Dim StartString As Long
    Dim EndString As Long
    Dim NumberIndex As Long
    Dim FechaTemp As String
    Dim NumberTemp As String
    
    On Error GoTo ErrorHandler
    
    'OBTENGO LOS DATOS DEL BUFFER
    BufferCurrent = comControl.Input
    If BufferCurrent = "" Then
        Exit Function
    End If
    
    Buffer = Buffer & BufferCurrent
    
    NumberIndex = InStr(1, Buffer, "RING")
    If NumberIndex > 0 Then
        RaiseEvent Ringing
    End If
    
    NumberIndex = InStr(1, Buffer, "NMBR = ")
    If NumberIndex > 0 Then
        RaiseEvent Identifying
        Do
            DoEvents
            Buffer = Buffer & comControl.Input
        Loop Until InStr(NumberIndex, Buffer, vbCrLf) > 0
        
        'DATE
        StartString = InStr(1, Buffer, "DATE = ") + 7
        FechaTemp = Year(Date) & "/" & Mid(Buffer, StartString, 2) & "/" & Mid(Buffer, StartString + 2, 2)
        
        'TIME
        StartString = InStr(1, Buffer, "TIME = ") + 7
        FechaTemp = FechaTemp & " " & Mid(Buffer, StartString, 2) & ":" & Mid(Buffer, StartString + 2, 2)
        
        mCallerIDDateTime = CDate(FechaTemp)
        
        'NUMBER
        StartString = InStr(1, Buffer, "NMBR = ") + 7
        EndString = InStr(StartString, Buffer, vbCrLf)
        NumberTemp = Trim(Mid(Buffer, StartString, EndString - StartString))
        
        Select Case NumberTemp
            Case ""
                mCallerIDNumber = CALLERID_NOTAVAILABLE
            Case "P"
                mCallerIDNumber = CALLERID_PRIVATE
            Case Else
                mCallerIDNumber = IIf(Left(NumberTemp, 1) <> "0", LocationAreaCode & NumberTemp, NumberTemp)
        End Select
        DoEvents
        
        Buffer = ""
        CallerIDCheck = True
    End If
    Exit Function
    
ErrorHandler:
End Function

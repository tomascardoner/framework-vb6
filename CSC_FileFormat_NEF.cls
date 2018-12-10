VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSC_FileFormat_NEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'PROPERTIES
Private mFilePathAndName As String
Private mFirmwareVersion As String
Private mISOAuto As Boolean
Private mISONumber As Integer
Private mColorMode As String
Private mQuality As String
Private mWhiteBalance As String
Private mSharpening As String
Private mFocusMode As String
Private mFlashSetting As String
Private mAutoFlashMode As String
Private mWhiteBalanceFine As Byte

Private mIsDirty As Boolean

Private Const HEADER_START_BYTE As Long = &H683

Public Property Get FilePathAndName() As String
    FilePathAndName = mFilePathAndName
End Property

Public Property Let FilePathAndName(ByVal Value As String)
    If mFilePathAndName <> Value Then
        mIsDirty = True
        mFilePathAndName = Value
    End If
End Property

Public Function LoadHeaderInfo() As Boolean
    Dim FileNumber As Integer
    Dim FileSize As Long
    Dim TempData As String
    Dim Offset As Long

    On Error GoTo ErroHandler
    
    If mFilePathAndName = "" Then
        Exit Function
    End If
    If Dir(mFilePathAndName) = "" Then
        Exit Function
    End If
    
    FileNumber = FreeFile()
    Open mFilePathAndName For Binary Access Read As FileNumber
    
    FileSize = LOF(FileNumber)
    If FileSize = 0 Then
        Close FileNumber
    End If
    
    Offset = HEADER_START_BYTE
    
    'Firmware Version
    TempData = String(4, vbNullChar)
    Get FileNumber, Offset, TempData
    'mFirmwareVersion
    Offset = Offset + Len(TempData)
    
    'ISO Setting (AUTO/MANUAL) And Number
    TempData = String(2, vbNullChar)
    Get FileNumber, Offset, TempData
    'mISO_Auto
    'mISO_Number
    Offset = Offset + Len(TempData)
    
    'Color Mode (COLOR)
    TempData = String(8, vbNullChar)
    Get FileNumber, Offset, TempData
    mColorMode = Trim(CleanNullChars(TempData))
    Offset = Offset + Len(TempData)
    
    'Quality (RAW, RAW2.7M)
    TempData = String(8, vbNullChar)
    Get FileNumber, Offset, TempData
    mQuality = Trim(CleanNullChars(TempData))
    Offset = Offset + Len(TempData)
    
    'White Balance (AUTO, CLOUDY, FLASH, INCANDESCENT, PRESET, PRESET1, SUNNY, WHITE PRESET)
    TempData = String(12, vbNullChar)
    Get FileNumber, Offset, TempData
    mQuality = Trim(CleanNullChars(TempData))
    Offset = Offset + Len(TempData)
    
    'Sharpening (AUTO, HIGH, LOW, MED.H, MED.L, NONE, NORMAL, empty)
    TempData = String(4, vbNullChar)
    Get FileNumber, Offset, TempData
    mSharpening = Trim(CleanNullChars(TempData))
    Offset = Offset + Len(TempData)
    
    'Focus Mode (AF-A, AF-C, AF-S, MANUAL)
    TempData = String(8, vbNullChar)
    Get FileNumber, Offset, TempData
    mFocusMode = Trim(CleanNullChars(TempData))
    Offset = Offset + Len(TempData)
    
    'Flash Setting (NORMAL, REAR, REAR SLOW, RED-EYES, LOW, empty)
    TempData = String(8, vbNullChar)
    Get FileNumber, Offset, TempData
    mFlashSetting = Trim(CleanNullChars(TempData))
    Offset = Offset + Len(TempData)
        
    LoadHeaderInfo = True
    Exit Function
    
ErroHandler:
End Function


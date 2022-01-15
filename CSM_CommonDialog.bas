Attribute VB_Name = "CSM_CommonDialog"
Option Explicit

'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'FILE OPEN
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
 
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'//////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'BROWSE FOLDER
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
'//////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'GET COLOR
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
 
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Private mCustomColorsInitialized As Boolean
Private mCustomColors() As Byte
'//////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'GET FONT
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
Private Const LF_FACESIZE As Long = 32

Private Const FW_BOLD As Long = 700
Private Const FW_HEAVY As Long = 900
Private Const FW_BLACK As Long = FW_HEAVY
Private Const FW_SEMIBOLD As Long = 600
Private Const FW_DEMIBOLD As Long = FW_SEMIBOLD
Private Const FW_DONTCARE As Long = 0
Private Const FW_EXTRABOLD As Long = 800
Private Const FW_EXTRALIGHT As Long = 200
Private Const FW_LIGHT As Long = 300
Private Const FW_MEDIUM As Long = 500
Private Const FW_NORMAL As Long = 400
Private Const FW_REGULAR As Long = FW_NORMAL
Private Const FW_THIN As Long = 100
Private Const FW_ULTRABOLD As Long = FW_EXTRABOLD
Private Const FW_ULTRALIGHT As Long = FW_EXTRALIGHT

Private Const DEFAULT_CHARSET As Long = 1
Private Const DEFAULT_PITCH As Long = 0
Private Const OUT_DEFAULT_PRECIS As Long = 0
Private Const CLIP_DEFAULT_PRECIS As Long = 0
Private Const DEFAULT_QUALITY As Long = 0
Private Const FF_DECORATIVE As Long = 80
Private Const FF_DONTCARE As Long = 0
Private Const FF_MODERN As Long = 48
Private Const FF_ROMAN As Long = 16
Private Const FF_SCRIPT As Long = 64
Private Const FF_SWISS As Long = 32
Private Const CF_SCREENFONTS As Long = &H1
Private Const CF_PRINTERFONTS As Long = &H2
Private Const CF_BOTH As Long = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS As Long = &H100&
Private Const CF_FORCEFONTEXIST As Long = &H10000
Private Const CF_INITTOLOGFONTSTRUCT As Long = &H40&
Private Const CF_LIMITSIZE As Long = &H2000&
Private Const REGULAR_FONTTYPE As Long = &H400

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * LF_FACESIZE
End Type

Private Type CHOOSEFONT
    lStructSize As Long
    hwndOwner As Long ' caller's window handle
    hdc As Long ' printer DC/IC or NULL
    lpLogFont As Long ' ptr. to a LOGFONT struct
    iPointSize As Long ' 10 size points of selected font
    flags As Long ' enum. type flags
    rgbColors As Long ' returned text color
    lCustData As Long ' data passed to hook fn.
    lpfnHook As Long ' ptr. to hook function
    lpTemplateName As String ' custom template name
    hInstance As Long ' stance handle of.EXE that
    ' contains cust. dlg. template
    lpszStyle As String ' return the style field here
    ' must be LF_FACESIZE or bigger
    nFontType As Integer ' same value reported to the EnumFonts
    ' call back with the extra FONTTYPE_
    ' bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long ' minimum pt size allowed &
    nSizeMax As Long ' max pt size allowed if
    ' CF_LIMITSIZE is used
End Type

Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function ChooseFontAPI Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
'//////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'PRINT
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
Private Declare Function PrintDlgA Lib "comdlg32.dll" (pPrintdlg As PRINTDLG) As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Const DM_ORIENTATION = &H1&
Private Const DM_DUPLEX = &H1000&
Private Const DM_COPIES = &H100&

Private Type PRINTDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

'//////////////////////////////////////////////////////////////////



'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'FILE OPEN
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////

Public Function FileOpen(ByVal hwndOwner As Long, ByVal Title As String, ByVal Filter As String, Optional ByVal InitialDir As String = "", Optional ByVal HideReadOnlyButton As Boolean = True, Optional ByVal AllowMultiSelect As Boolean = False, Optional ByVal ExplorerStyle As Boolean = True, Optional ByVal FileMustExist As Boolean = True) As String
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim flags As Long
    
    Const OFN_HIDEREADONLY = &H4
    Const OFN_ALLOWMULTISELECT = &H200
    Const OFN_EXPLORER = &H80000
    Const OFN_FILEMUSTEXIST = &H1000
    
    Filter = Replace(Filter, "|", vbNullChar, 1)
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = hwndOwner
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = Filter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(255, vbNullChar)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = InitialDir
    OpenFile.lpstrTitle = Title
    If HideReadOnlyButton Then
        flags = flags Or OFN_HIDEREADONLY
    End If
    If AllowMultiSelect Then
        flags = flags Or OFN_ALLOWMULTISELECT
    End If
    If ExplorerStyle Then
        flags = flags Or OFN_EXPLORER
    End If
    If FileMustExist Then
        flags = flags Or OFN_FILEMUSTEXIST
    End If
    OpenFile.flags = flags
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
       FileOpen = ""
    Else
       FileOpen = Trim(OpenFile.lpstrFile)
    End If
End Function

'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'BROWSE FOR FOLDER
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////

Public Function BrowseForFolder(ByVal hwndOwner As Long, ByVal Title As String) As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
 
    With tBrowseInfo
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(Title, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    End If
End Function


'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'GET COLOR
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////

Private Sub InitializeCustomColors()
    ReDim mCustomColors(0 To (16 * 4) - 1) As Byte
    
    Dim i As Integer
    
    For i = LBound(mCustomColors) To UBound(mCustomColors) Step 4
        mCustomColors(i) = 255
        mCustomColors(i + 1) = 255
        mCustomColors(i + 2) = 255
        mCustomColors(i + 3) = 0
    Next i
    
    mCustomColorsInitialized = True
End Sub

Public Function GetColor(ByVal hwndOwner As Long, ByRef Color As Long) As Boolean
    Dim cc As CHOOSECOLOR
    Dim lReturn As Long
    
    If Not mCustomColorsInitialized Then
        Call InitializeCustomColors
    End If
    
    cc.lStructSize = Len(cc)
    cc.hwndOwner = hwndOwner
    cc.hInstance = 0
    cc.lpCustColors = StrConv(mCustomColors, vbUnicode)
    cc.flags = 0
    lReturn = ChooseColorAPI(cc)
    If lReturn <> 0 Then
        Color = cc.rgbResult
        mCustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
        GetColor = True
    Else
        GetColor = False
    End If
End Function

'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'GET FONT
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////

Public Function GetFont(ByVal hwndOwner As Long, ByRef FontName As String, ByRef FontBold As Boolean, ByRef FontItalic As Boolean, ByRef FontUnderline As Boolean, ByRef FontStrikethru As Boolean, ByRef FontSize As Single, ByVal AllowSelectColor As Boolean, ByRef FontColor As Long) As Boolean
    Dim cf As CHOOSEFONT
    Dim lfont As LOGFONT
    Dim hMem As Long
    Dim pMem As Long
    Dim retval As Long
    
    lfont.lfHeight = FontSize * 96 / 72      ' determine default height
    lfont.lfWidth = 0 ' determine default width
    lfont.lfEscapement = 0  ' angle between baseline and escapement vector
    lfont.lfOrientation = 0  ' angle between baseline and orientation vector
    lfont.lfWeight = IIf(FontBold, FW_BOLD, FW_NORMAL)
    lfont.lfItalic = Abs(FontItalic)
    lfont.lfUnderline = Abs(FontUnderline)
    lfont.lfStrikeOut = Abs(FontStrikethru)
    lfont.lfCharSet = DEFAULT_CHARSET  ' use default character set
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
    lfont.lfQuality = DEFAULT_QUALITY  ' default quality setting
    lfont.lfPitchAndFamily = DEFAULT_PITCH 'Or FF_ROMAN  ' default pitch, proportional with serifs
    lfont.lfFaceName = FontName & vbNullChar  ' string must be null-terminated
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)  ' lock and get pointer
    CopyMemory ByVal pMem, lfont, Len(lfont)  ' copy structure's contents into block
    ' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
    cf.lStructSize = Len(cf)  ' size of structure
    cf.hwndOwner = hwndOwner  ' window Form1 is opening this dialog box
    cf.hdc = Printer.hdc  ' device context of default printer (using VB's mechanism)
    cf.lpLogFont = pMem   ' pointer to LOGFONT memory block buffer
    cf.iPointSize = 0   ' 12 point font (in units of 1/10 point)
    If AllowSelectColor Then
        cf.flags = CF_SCREENFONTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE Or CF_EFFECTS
    Else
        cf.flags = CF_SCREENFONTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    End If
    cf.rgbColors = FontColor  ' black
    cf.nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
    cf.nSizeMin = 6  ' minimum point size
    cf.nSizeMax = 72  ' maximum point size
    ' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
    ' and then print out the attributes we mentioned earlier that the user selected.
    retval = ChooseFontAPI(cf)  ' open the dialog box
    If retval <> 0 Then  ' success
        CopyMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
        ' Now make the fixed-length string holding the font name into a "normal" string.
        FontName = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        FontBold = (lfont.lfWeight = FW_BOLD)
        FontItalic = CBool(lfont.lfItalic)
        FontUnderline = CBool(lfont.lfUnderline)
        FontStrikethru = CBool(lfont.lfStrikeOut)
        FontSize = cf.iPointSize / 10
        FontColor = cf.rgbColors
        GetFont = True
    End If
    ' Deallocate the memory block we created earlier.  Note that this must
    ' be done whether the function succeeded or not.
    retval = GlobalUnlock(hMem)  ' destroy pointer, unlock block
    retval = GlobalFree(hMem)  ' free the allocated memory
End Function


'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
'PRINT
'//////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////
Public Function ShowPrint(ByVal hwndOwner As Long, Optional PrintFlags As Long) As Boolean
    Dim PRINTDLG As PRINTDLG
    Dim DEVMODE As DEVMODE
    Dim DevName As DEVNAMES
    
    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String
    Dim strSetting As String
    
    ' Use PrintSetupDialog to get the handle to a memory
    ' block with a DevMode and DevName structures
    
    PRINTDLG.lStructSize = Len(PRINTDLG)
    PRINTDLG.hwndOwner = hwndOwner
    
    PRINTDLG.flags = PrintFlags
    
    ' Set the current orientation and duplex setting
    DEVMODE.dmDeviceName = Printer.DeviceName
    DEVMODE.dmSize = Len(DEVMODE)
    DEVMODE.dmFields = DM_ORIENTATION Or DM_DUPLEX Or DM_COPIES
    DEVMODE.dmOrientation = Printer.Orientation
    DEVMODE.dmCopies = Printer.Copies
    On Error Resume Next
    DEVMODE.dmDuplex = Printer.Duplex
    On Error GoTo 0
    
    ' Allocate memory for the initialization hDevMode structure
    ' and copy the settings gathered above into this memory
    PRINTDLG.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DEVMODE))
    lpDevMode = GlobalLock(PRINTDLG.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DEVMODE, Len(DEVMODE)
        bReturn = GlobalUnlock(PRINTDLG.hDevMode)
    End If
    
    ' Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    
    ' Allocate memory for the initial hDevName structure
    ' and copy the settings gathered above into this memory
    PRINTDLG.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PRINTDLG.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If
    
    ' Call the print dialog up and let the user make changes
    If PrintDlgA(PRINTDLG) Then
    
        ' First get the DevName structure.
        lpDevName = GlobalLock(PRINTDLG.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PRINTDLG.hDevNames
        
        ' Next get the DevMode structure and set the printer
        ' properties appropriately
        lpDevMode = GlobalLock(PRINTDLG.hDevMode)
        CopyMemory DEVMODE, ByVal lpDevMode, Len(DEVMODE)
        bReturn = GlobalUnlock(PRINTDLG.hDevMode)
        GlobalFree PRINTDLG.hDevMode
        NewPrinterName = UCase$(Left(DEVMODE.dmDeviceName, InStr(DEVMODE.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                End If
            Next
        End If
        On Error Resume Next
        
        ' Set printer object properties according to selections made
        ' by user
        DoEvents
        With Printer
            .Copies = DEVMODE.dmCopies
            .Duplex = DEVMODE.dmDuplex
            .Orientation = DEVMODE.dmOrientation
        End With
        On Error GoTo 0
        ShowPrint = True
    End If
    
    ' Display the results in the immediate (debug) window
    With Printer
        If .Orientation = 1 Then
            strSetting = "Portrait.  "
        Else
            strSetting = "Landscape. "
        End If
    End With
End Function


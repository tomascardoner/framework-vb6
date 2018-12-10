Attribute VB_Name = "CSM_Forms"
Option Explicit

Private Const STATE_SYSTEM_FOCUSABLE = &H100000
Private Const STATE_SYSTEM_INVISIBLE = &H8000
Private Const STATE_SYSTEM_OFFSCREEN = &H10000
Private Const STATE_SYSTEM_UNAVAILABLE = &H1
Private Const STATE_SYSTEM_PRESSED = &H8
Private Const CCHILDREN_TITLEBAR = 5

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TITLEBARINFO
    cbSize As Long
    rcTitleBar As RECT
    rgstate(CCHILDREN_TITLEBAR) As Long
End Type

Private Declare Function GetTitleBarInfo Lib "user32.dll" (ByVal hwnd As Long, ByRef pti As TITLEBARINFO) As Long

Public Sub SetScaleHeight(ByRef TheForm As Form, ByVal NewScaleHeight As Single)
    Dim TitleInfo As TITLEBARINFO
    Dim WindowTitleHeight As Long
    Const WINDOW_BORDER_SIZE_APROX As Long = 2
    
    'Initialize structure
    TitleInfo.cbSize = Len(TitleInfo)
    'Retrieve information about the tilte bar of this window
    GetTitleBarInfo TheForm.hwnd, TitleInfo
    WindowTitleHeight = (TitleInfo.rcTitleBar.Bottom - TitleInfo.rcTitleBar.Top) * Screen.TwipsPerPixelY
    
    TheForm.AutoRedraw = False
    
    If TheForm.ScaleHeight <> NewScaleHeight Then
        TheForm.Height = NewScaleHeight + WindowTitleHeight + (WINDOW_BORDER_SIZE_APROX * 2)
    End If
    Do While TheForm.ScaleHeight <> NewScaleHeight
        Select Case TheForm.ScaleHeight - NewScaleHeight
            
            'MENOR
            Case Is <= -500
                TheForm.Height = TheForm.Height + 500
            Case Is <= -200
                TheForm.Height = TheForm.Height + 200
            Case Is <= -100
                TheForm.Height = TheForm.Height + 100
            Case Is <= -50
                TheForm.Height = TheForm.Height + 50
            Case Is <= -20
                TheForm.Height = TheForm.Height + 20
            Case Is <= -10
                TheForm.Height = TheForm.Height + 10
            Case Is <= -5
                TheForm.Height = TheForm.Height + 5
            Case Is <= -1
                TheForm.Height = TheForm.Height + 1
            
            'MAYOR
            Case Is >= 500
                TheForm.Height = TheForm.Height - 500
            Case Is <= 200
                TheForm.Height = TheForm.Height - 200
            Case Is <= 100
                TheForm.Height = TheForm.Height - 100
            Case Is <= 50
                TheForm.Height = TheForm.Height - 50
            Case Is <= 20
                TheForm.Height = TheForm.Height - 20
            Case Is <= 10
                TheForm.Height = TheForm.Height - 10
            Case Is <= 5
                TheForm.Height = TheForm.Height - 5
            Case Is <= 1
                TheForm.Height = TheForm.Height - 1
            
        End Select
    Loop
    
    TheForm.AutoRedraw = True
End Sub

Public Function IsLoaded(ByVal FormName As String) As Boolean
    Dim frmCurrent As VB.Form
    
    For Each frmCurrent In VB.Forms
        If frmCurrent.Name = FormName Then
            IsLoaded = True
            Exit Function
        End If
    Next frmCurrent
    IsLoaded = False
End Function

Public Function GetLoaded(ByVal FormName As String) As Form
    Dim frmCurrent As VB.Form
    
    For Each frmCurrent In VB.Forms
        If frmCurrent.Name = FormName Then
            Set GetLoaded = frmCurrent
        End If
    Next frmCurrent
End Function

Public Sub ResizeAndPosition(ByRef MDIForm As MDIForm, ByRef MDIChildForm As Form)
    On Error Resume Next

    With MDIChildForm
        If .BorderStyle = vbSizable Then
            .Top = 0
            .Left = 0
            .Height = MDIForm.ScaleHeight
            .Width = MDIForm.ScaleWidth
        End If
    End With
End Sub

Public Sub ResizeAndPositionGeneric(ByRef MDIForm As MDIForm, ByRef MDIChildForm As Form)
    On Error Resume Next

    With MDIChildForm
        If .BorderStyle = vbSizable Then
            .Top = 0
            .Left = 0
            .Height = MDIForm.ScaleHeight
            .Width = MDIForm.ScaleWidth
        End If
    End With
End Sub

Public Sub ResizeAndPositionAll(ByRef MDIForm As MDIForm)
    Dim frmCurrent As Form
    
    Screen.MousePointer = vbHourglass
    
    'Para cada Form cargado
    For Each frmCurrent In VB.Forms
        If frmCurrent.Name <> MDIForm.Name Then
            Call ResizeAndPosition(MDIForm, frmCurrent)
        End If
    Next frmCurrent
    
    Screen.MousePointer = vbDefault
End Sub

Public Function UnloadAll(ParamArray avExceptForms() As Variant)
    Dim frmCurrent As VB.Form
    Dim vntFormName As Variant
    Dim blnExceptCurrent As Boolean
    
    Screen.MousePointer = vbHourglass
    
    '23-Jan-2002: Check if every item in ParamArray are string
    For Each vntFormName In avExceptForms
        If VarType(vntFormName) <> vbString Then
            If pIsCompiled Then
                CSM_Error.ShowErrorMessage "UnloadForms", "The ParamArray Element is not a string.", False
            Else
                Stop
            End If
        End If
    Next vntFormName
    
    For Each frmCurrent In VB.Forms
    
        blnExceptCurrent = False
        
        'Para cada Nombre de Form en el Parameter Array...
        For Each vntFormName In avExceptForms
            'Si el form que estoy por descargar, está en el Parameter Array...
            If frmCurrent.Name = vntFormName Then
                blnExceptCurrent = True
                Exit For
            End If
        Next vntFormName
        
        'Si no es una excepción...
        If Not blnExceptCurrent Then
            Unload frmCurrent
            Set frmCurrent = Nothing
        End If
    Next frmCurrent
    
    Screen.MousePointer = vbDefault
End Function

Public Sub CenterToParent(ByRef ParentForm As Form, ByRef ChildForm As Form)
    On Error Resume Next
    
    With ParentForm
        If ChildForm.MDIChild Then
            ChildForm.Top = ((.ScaleHeight - ChildForm.Height) / 2)
            ChildForm.Left = ((.ScaleWidth - ChildForm.Width) / 2)
        Else
            ChildForm.Top = .Top + ((.Height - ChildForm.Height) / 2)
            ChildForm.Left = .Left + ((.Width - ChildForm.Width) / 2)
        End If
    End With
End Sub

Public Function GetProperties(ByRef TheForm As Form) As CSC_FormProperties
    Dim FormProperties As CSC_FormProperties
    
    Set FormProperties = New CSC_FormProperties
    
    With FormProperties
        .Top = TheForm.Top
        .Left = TheForm.Left
        .Height = TheForm.Height
        .Width = TheForm.Width
    End With
    
    Set GetProperties = FormProperties
End Function

Public Function GetFormIndex(ByVal FormName As String) As Long
    Dim FormIndex As Long
    
    If FormName <> "" Then
        For FormIndex = 0 To Forms.Count - 1
            If Forms(FormIndex).Name = FormName Then
                GetFormIndex = FormIndex
                Exit Function
            End If
        Next FormIndex
    End If
    GetFormIndex = -1
End Function

Public Function ControlsChangeEnabledState(ByRef TheForm As Form, ByVal NewEnabledState As Boolean, ByVal ApplyToLabels As Boolean, ByVal ApplyToFrames As Boolean, ParamArray ControlsExcept() As Variant)
    Dim ControlCurrent As Control
    Dim vntControlName As Variant
    Dim blnExceptCurrent As Boolean
    
    For Each ControlCurrent In TheForm.Controls
        blnExceptCurrent = False
        For Each vntControlName In ControlsExcept
            If CStr(vntControlName) = CStr(ControlCurrent.Name) Then
                blnExceptCurrent = True
                Exit For
            End If
        Next vntControlName
        
        If Not blnExceptCurrent Then
            Select Case TypeName(ControlCurrent)
                Case "Label"
                    If ApplyToLabels Then
                        On Error Resume Next
                        ControlCurrent.Enabled = NewEnabledState
                    End If
                Case "Frame"
                    If ApplyToFrames Then
                        On Error Resume Next
                        ControlCurrent.Enabled = NewEnabledState
                    End If
                Case Else
                    On Error Resume Next
                    ControlCurrent.Enabled = NewEnabledState
            End Select
        End If
    Next ControlCurrent
End Function

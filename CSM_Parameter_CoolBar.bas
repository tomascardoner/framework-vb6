Attribute VB_Name = "CSM_Parameter_CoolBar"
Option Explicit

Public Sub GetSettings(ByVal Name As String, ByRef CoolBar As ComCtl3.CoolBar)
    Dim Band As ComCtl3.Band
    
    On Error GoTo ErrorHandler
    
    For Each Band In CoolBar.Bands
        Band.Width = CSM_Registry.GetValue_FromApplication_LocalMachine("Interface\" & Name, "CoolBar_Band_" & Band.Key & "_Width", Band.Width, csrdtNumberInteger)
        If Band.Position > 1 Then
            Band.NewRow = CSM_Registry.GetValue_FromApplication_LocalMachine("Interface\" & Name, "CoolBar_Band_" & Band.Key & "_NewRow", Band.NewRow, csrdtBoolean)
        End If
    Next Band
    
ErrorHandler:
End Sub

Public Sub GetSettingsFromTag(ByRef CoolBar As ComCtl3.CoolBar)
    Dim Band As ComCtl3.Band

    On Error GoTo ErrorHandler

    For Each Band In CoolBar.Bands
        If Band.Position > 1 Then
            Band.NewRow = CBool(CSM_String.GetSubString(Band.Tag, 1, KEY_DELIMITER))
        End If
        Band.Width = CSng(CSM_String.GetSubString(Band.Tag, 2, KEY_DELIMITER))
    Next Band

ErrorHandler:
End Sub

Public Sub DeleteSettings(ByVal Name As String, ByRef CoolBar As ComCtl3.CoolBar)
    Dim Band As ComCtl3.Band

    On Error GoTo ErrorHandler

    DoEvents
    For Each Band In CoolBar.Bands
        Call CSM_Registry.DeleteValue_FromApplication_LocalMachine("Interface\" & Name, "CoolBar_Band_" & Band.Key & "_Width")
        If Band.Position > 1 Then
            Call CSM_Registry.DeleteValue_FromApplication_LocalMachine("Interface\" & Name, "CoolBar_Band_" & Band.Key & "_NewRow")
        End If
    Next Band

ErrorHandler:
End Sub

Public Sub SaveSettings(ByVal Name As String, ByRef CoolBar As ComCtl3.CoolBar)
    Dim Band As ComCtl3.Band
    
    On Error GoTo ErrorHandler
    
    DoEvents
    For Each Band In CoolBar.Bands
        Call CSM_Registry.SetValue_ToApplication_LocalMachine("Interface\" & Name, "CoolBar_Band_" & Band.Key & "_Width", Band.Width)
        If Band.Position > 1 Then
            Call CSM_Registry.SetValue_ToApplication_LocalMachine("Interface\" & Name, "CoolBar_Band_" & Band.Key & "_NewRow", Band.NewRow)
        End If
    Next Band
    
ErrorHandler:
End Sub

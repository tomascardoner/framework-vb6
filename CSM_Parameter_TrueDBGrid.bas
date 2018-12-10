Attribute VB_Name = "CSM_Parameter_TrueDBGrid"
Option Explicit

Public Sub GetSettingsWidth(ByVal Name As String, ByRef TrueDBGrid As TrueOleDBGrid80.TDBGrid)
    Dim Column As TrueOleDBGrid80.Column

    For Each Column In TrueDBGrid.Columns
        Column.Width = CSM_Registry.GetValue_FromApplication_LocalMachine("Interface\" & Name, "Grid_Column_" & Column.DataField & "_Width", Column.Width, csrdtNumberDecimal)
    Next Column
End Sub

Public Sub GetSettingsPosition(ByVal Name As String, ByRef TrueDBGrid As TrueOleDBGrid80.TDBGrid)
    Dim Column As TrueOleDBGrid80.Column
    
    On Error Resume Next

    For Each Column In TrueDBGrid.Columns
        Column.Order = CSM_Registry.GetValue_FromApplication_LocalMachine("Interface\" & Name, "Grid_Column_" & Column.DataField & "_Position", Column.Order, csrdtNumberInteger)
    Next Column
End Sub

Public Sub GetSettings(ByVal Name As String, ByRef TrueDBGrid As Object)
    Call GetSettingsWidth(Name, TrueDBGrid)
    Call GetSettingsPosition(Name, TrueDBGrid)
End Sub

Public Sub DeleteSettingsWidth(ByVal Name As String, ByRef TrueDBGrid As TrueOleDBGrid80.TDBGrid)
    Dim Column As TrueOleDBGrid80.Column

    For Each Column In TrueDBGrid.Columns
        Call CSM_Registry.DeleteValue_FromApplication_LocalMachine("Interface\" & Name, "Grid_Column_" & Column.DataField & "_Width")
    Next Column
End Sub

Public Sub DeleteSettingsPosition(ByVal Name As String, ByRef TrueDBGrid As TrueOleDBGrid80.TDBGrid)
    Dim Column As TrueOleDBGrid80.Column

    For Each Column In TrueDBGrid.Columns
        Call CSM_Registry.DeleteValue_FromApplication_LocalMachine("Interface\" & Name, "Grid_Column_" & Column.DataField & "_Position")
    Next Column
End Sub

Public Sub SaveSettings(ByVal Name As String, ByRef TrueDBGrid As Object)
    Dim Column As Object
    
    For Each Column In TrueDBGrid.Columns
        Call CSM_Registry.SetValue_ToApplication_LocalMachine("Interface\" & Name, "Grid_Column_" & Column.DataField & "_Width", Column.Width)
        Call CSM_Registry.SetValue_ToApplication_LocalMachine("Interface\" & Name, "Grid_Column_" & Column.DataField & "_Position", Column.Order)
    Next Column
End Sub

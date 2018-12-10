Attribute VB_Name = "CSM_Parameter_ListView"
Option Explicit

Public Sub GetSettings(ByVal Name As String, ByRef ListView As MSComctlLib.ListView)
    Dim ColumnHeader As MSComctlLib.ColumnHeader
    
    On Error GoTo ErrorHandler
    
    ListView.SortKey = CSM_Registry.GetValue_FromApplication_LocalMachine("Interface\" & Name, "ListView_SortKey", ListView.SortKey, csrdtNumberInteger)
    ListView.SortOrder = CSM_Registry.GetValue_FromApplication_LocalMachine("Interface\" & Name, "ListView_SortOrder", ListView.SortOrder, csrdtNumberInteger)
    For Each ColumnHeader In ListView.ColumnHeaders
        ColumnHeader.Width = CSM_Registry.GetValue_FromApplication_LocalMachine("Interface\" & Name, "ListView_Column_" & ColumnHeader.Key & "_Width", ColumnHeader.Width, csrdtNumberDecimal)
        ColumnHeader.Position = CSM_Registry.GetValue_FromApplication_LocalMachine("Interface\" & Name, "ListView_Column_" & ColumnHeader.Key & "_Position", ColumnHeader.Position, csrdtNumberInteger)
    Next ColumnHeader
    
ErrorHandler:
End Sub

Public Sub SaveSettings(ByVal Name As String, ByRef ListView As MSComctlLib.ListView)
    Dim ColumnHeader As MSComctlLib.ColumnHeader
    
    On Error GoTo ErrorHandler
    
    DoEvents
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Interface\" & Name, "ListView_SortKey", ListView.SortKey)
    Call CSM_Registry.SetValue_ToApplication_LocalMachine("Interface\" & Name, "ListView_SortOrder", ListView.SortOrder)
    For Each ColumnHeader In ListView.ColumnHeaders
        Call CSM_Registry.SetValue_ToApplication_LocalMachine("Interface\" & Name, "ListView_Column_" & ColumnHeader.Key & "_Width", ColumnHeader.Width)
        Call CSM_Registry.SetValue_ToApplication_LocalMachine("Interface\" & Name, "ListView_Column_" & ColumnHeader.Key & "_Position", ColumnHeader.Position)
    Next ColumnHeader
    
ErrorHandler:
End Sub

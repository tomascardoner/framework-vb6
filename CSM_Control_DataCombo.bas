Attribute VB_Name = "CSM_Control_DataCombo"
Option Explicit

Public Function FillFromSQL(ByRef DataCombo As MSDataListLib.DataCombo, ByVal RecordSource As String, ByVal BoundField As String, ByVal ListField As String, ByVal ErrorEntityName As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Variant = Empty) As Boolean
    Dim MousePointerSave As Integer
    Dim cmdData As ADODB.command
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Set cmdData = New ADODB.command
    Set cmdData.ActiveConnection = pDatabase.Connection
    cmdData.CommandText = RecordSource
    cmdData.CommandType = adCmdText
    
    FillFromSQL = FillFromCommand(DataCombo, cmdData, BoundField, ListField, ErrorEntityName, ItemPosition, ItemValue)
    
    Set cmdData = Nothing
    
    Screen.MousePointer = MousePointerSave
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Modules.CSM_Control_DataCombo.FillFromSQL", "Error al leer la lista de " & ErrorEntityName & "."
End Function

Public Function FillFromCommand(ByRef DataCombo As MSDataListLib.DataCombo, ByRef command As ADODB.command, ByVal BoundField As String, ByVal ListField As String, ByVal ErrorEntityName As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Variant = Empty) As Boolean
    Dim MousePointerSave As Integer
    Dim recData As ADODB.Recordset
    Dim SelectedValue As String
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    SelectedValue = DataCombo.BoundText
    
    DataCombo.BoundColumn = BoundField
    DataCombo.ListField = ListField
    
    Set recData = New ADODB.Recordset
    recData.Open command, , adOpenStatic, adLockReadOnly
    
    Set DataCombo.RowSource = recData
    DataCombo.BoundText = SelectedValue
    
    Call FindItemInternal(DataCombo, SelectedValue, ErrorEntityName, ItemPosition, ItemValue)
    
    Set recData = Nothing
    Set command = Nothing
    
    Screen.MousePointer = MousePointerSave
    FillFromCommand = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Modules.CSM_Control_DataCombo.FillFromCommand", "Error al leer la lista de " & ErrorEntityName & "."
End Function

Private Sub FindItemInternal(ByRef DataCombo As MSDataListLib.DataCombo, ByVal SelectedValue As String, ByVal ErrorEntityName As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Variant = Empty)
    Dim MousePointerSave As Integer
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If DataCombo.RowSource Is Nothing Then
        Exit Sub
    End If
    Set recData = DataCombo.RowSource
    
    Select Case ItemPosition
        Case cscpNone
            DataCombo.BoundText = ""
        Case cscpFirst
            If Not recData.EOF Then
                DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
            End If
        Case cscpFirstIfUnique
            If Not recData.EOF Then
                If recData.RecordCount = 1 Then
                    DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
                End If
            End If
        Case cscpLast
            If Not recData.EOF Then
                recData.MoveLast
                DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
            End If
        Case cscpCurrentOrNone
            If Not recData.EOF Then
                DataCombo.BoundText = SelectedValue
            End If
        Case cscpCurrentOrFirst
            If Not recData.EOF Then
                DataCombo.BoundText = SelectedValue
                If SelectedValue = "" Or DataCombo.BoundText <> SelectedValue Then
                    DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
                End If
            End If
        Case cscpCurrentOrFirstIfUnique
            If Not recData.EOF Then
                DataCombo.BoundText = SelectedValue
                If (SelectedValue = "" Or DataCombo.BoundText <> SelectedValue) And recData.RecordCount = 1 Then
                    DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
                End If
            End If
        Case cscpCurrentOrLast
            If Not recData.EOF Then
                DataCombo.BoundText = SelectedValue
                If DataCombo.BoundText <> SelectedValue Then
                    recData.MoveLast
                    DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
                End If
            End If
        Case cscpItemOrNone
            DataCombo.BoundText = ItemValue
        Case cscpItemOrFirst
            If Not recData.EOF Then
                DataCombo.BoundText = ItemValue
                If (ItemValue = "" Or DataCombo.BoundText <> ItemValue) Then
                    DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
                End If
            End If
        Case cscpItemOrFirstIfUnique
            If Not recData.EOF Then
                DataCombo.BoundText = ItemValue
                If DataCombo.BoundText <> ItemValue And recData.RecordCount = 1 Then
                    DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
                End If
            End If
        Case cscpItemOrLast
            DataCombo.BoundText = ItemValue
            If DataCombo.BoundText <> ItemValue And Not recData.EOF Then
                recData.MoveLast
                DataCombo.BoundText = recData(DataCombo.BoundColumn).Value
            End If
    End Select

    Set recData = Nothing
    
    Screen.MousePointer = MousePointerSave
    Exit Sub
    
ErrorHandler:
    ShowErrorMessage "Modules.CSM_Control_DataCombo.FindItemInternal", "Error al buscar el item en la lista de " & ErrorEntityName & "."
End Sub

Public Sub FindItem(ByRef DataCombo As MSDataListLib.DataCombo, ByVal ErrorEntityName As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Variant = Empty)
    Dim SelectedValue As String
    
    SelectedValue = DataCombo.BoundText
    
    Call FindItemInternal(DataCombo, SelectedValue, ErrorEntityName, ItemPosition, ItemValue)
End Sub

Public Function GetSubID(ByRef DataCombo As MSDataListLib.DataCombo, ByVal SubIDLen As Byte, ByVal SubIDToGet As Byte) As Long
    Dim ComboString As String
    
    If DataCombo.BoundText <> "" Then
        ComboString = CStr(DataCombo.BoundText)
        
        Select Case SubIDToGet
            Case 1
                GetSubID = Val(Left(ComboString, Len(ComboString) - SubIDLen))
            Case 2
                GetSubID = Val(Right(ComboString, SubIDLen))
            Case Else
        End Select
    End If
End Function

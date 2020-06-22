Attribute VB_Name = "CSM_Control_ComboBox"
Option Explicit

Public Enum csComboPosition
    cscpNone = 1
    cscpFirst = 2
    cscpFirstIfUnique = 3
    cscpLast = 4
    cscpCurrentOrNone = 5
    cscpCurrentOrFirst = 6
    cscpCurrentOrFirstIfUnique = 7
    cscpCurrentOrFirstIfUniqueIgnoringPreloaded = 8
    cscpCurrentOrLast = 9
    cscpItemOrNone = 10
    cscpItemOrFirst = 11
    cscpItemOrFirstIfUnique = 12
    cscpItemOrFirstIfUniqueIgnoringPreloaded = 13
    cscpItemOrLast = 14
End Enum

Public Sub SelAllText(ByRef TheControl As ComboBox)
    On Error Resume Next
    
    TheControl.SelStart = 0
    TheControl.SelLength = Len(TheControl.Text)
End Sub

Public Function FillFromSQL(ByRef ComboBox As VB.ComboBox, ByVal RecordSource As String, ByVal BoundField As String, ByVal ListField As String, ByVal ErrorEntityName As String, Optional ByVal ItemPosition As csComboPosition = cscpNone, Optional ByVal ItemValue As Long = 0, Optional ByVal ClearComboBox As Boolean = True) As Boolean
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
    
    FillFromSQL = FillFromCommand(ComboBox, cmdData, BoundField, ListField, ErrorEntityName, ItemPosition, ItemValue, ClearComboBox)
    
    Set cmdData = Nothing
    
    Screen.MousePointer = MousePointerSave
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Modules.CSM_Control_ComboBox.FillFromSQL", "Error al leer la lista de " & ErrorEntityName & "."
End Function

Public Function FillFromCommand(ByRef ComboBox As VB.ComboBox, ByRef command As ADODB.command, ByVal BoundField As String, ByVal ListField As String, ByVal ErrorEntityName As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Long = 0, Optional ClearComboBox As Boolean = True) As Boolean
    Dim MousePointerSave As Integer
    Dim recData As ADODB.Recordset
    Dim CurrentValue As Long
    Dim PreloadedItemCount As Long
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If ComboBox.ListIndex > -1 Then
        CurrentValue = ComboBox.ItemData(ComboBox.ListIndex)
    End If
    
    If ClearComboBox Then
        ComboBox.Clear
    Else
        PreloadedItemCount = ComboBox.ListCount
    End If
    
    Set recData = New ADODB.Recordset
    recData.Open command, , adOpenForwardOnly, adLockReadOnly
    
    Do While Not recData.EOF
        ComboBox.AddItem recData(ListField).Value & ""
        If BoundField <> "" Then
            ComboBox.ItemData(ComboBox.NewIndex) = recData(BoundField).Value
        End If
        recData.MoveNext
    Loop
    
    Select Case ItemPosition
        Case cscpNone
            ComboBox.ListIndex = -1
        Case cscpFirst
            If ComboBox.ListCount > 0 Then
                ComboBox.ListIndex = 0
            End If
        Case cscpFirstIfUnique
            If ComboBox.ListCount = 1 Then
                ComboBox.ListIndex = 0
            Else
                ComboBox.ListIndex = -1
            End If
        Case cscpLast
            If ComboBox.ListCount > 0 Then
                ComboBox.ListIndex = ComboBox.ListCount - 1
            End If
        Case cscpItemOrNone, cscpItemOrfirst, cscpItemOrFirstIfUnique, cscpItemOrFirstIfUniqueIgnoringPreloaded, cscpItemOrLast
            ComboBox.ListIndex = GetListIndexByItemData(ComboBox, ItemValue, ItemPosition, PreloadedItemCount)
        Case cscpCurrentOrNone, cscpCurrentOrFirst, cscpCurrentOrFirstIfUnique, cscpCurrentOrFirstIfUniqueIgnoringPreloaded, cscpCurrentOrLast
            ComboBox.ListIndex = GetListIndexByItemData(ComboBox, CurrentValue, ItemPosition, PreloadedItemCount)
    End Select
    
    Set recData = Nothing
    Set command = Nothing
    
    Screen.MousePointer = MousePointerSave
    FillFromCommand = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Modules.CSM_Control_ComboBox.FillFromCommand", "Error al leer la lista de " & ErrorEntityName & "."
End Function

Public Function FillFromDelimitedString(ByRef ComboBox As VB.ComboBox, ByVal DelimitedString As String, ByVal Delimiter As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Long = 0, Optional ClearComboBox As Boolean = True) As Boolean
    Dim aValues() As String
    Dim Index As Long
    Dim CurrentValue As Long
    Dim PreloadedItemCount As Long
    
    If ComboBox.ListIndex > -1 Then
        CurrentValue = ComboBox.ItemData(ComboBox.ListIndex)
    End If
    
    If ClearComboBox Then
        ComboBox.Clear
    Else
        PreloadedItemCount = ComboBox.ListCount
    End If
    
    If Len(DelimitedString) > 0 Then
        aValues() = Split(DelimitedString, Delimiter)
        For Index = 0 To UBound(aValues)
            ComboBox.AddItem aValues(Index)
        Next Index
    End If
    
    Select Case ItemPosition
        Case cscpNone
            ComboBox.ListIndex = -1
        Case cscpFirst
            If ComboBox.ListCount > 0 Then
                ComboBox.ListIndex = 0
            End If
        Case cscpFirstIfUnique
            If ComboBox.ListCount = 1 Then
                ComboBox.ListIndex = 0
            Else
                ComboBox.ListIndex = -1
            End If
        Case cscpLast
            If ComboBox.ListCount > 0 Then
                ComboBox.ListIndex = ComboBox.ListCount - 1
            End If
        Case cscpItemOrNone, cscpItemOrfirst, cscpItemOrFirstIfUnique, cscpItemOrFirstIfUniqueIgnoringPreloaded, cscpItemOrLast
            ComboBox.ListIndex = GetListIndexByItemData(ComboBox, ItemValue, ItemPosition, PreloadedItemCount)
        Case cscpCurrentOrNone, cscpCurrentOrFirst, cscpCurrentOrFirstIfUnique, cscpCurrentOrFirstIfUniqueIgnoringPreloaded, cscpCurrentOrLast
            ComboBox.ListIndex = GetListIndexByItemData(ComboBox, CurrentValue, ItemPosition, PreloadedItemCount)
    End Select
    
    FillFromDelimitedString = True
End Function

Public Function FillFromCollection(ByRef ComboBox As VB.ComboBox, ByRef Collection As Collection, ByVal IDPropertyName As String, ByVal TextPropertyName As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Long = 0, Optional ClearComboBox As Boolean = True) As Boolean
    Dim CollectionItem As Variant
    Dim CurrentValue As Long
    Dim PreloadedItemCount As Long
    
    If ComboBox.ListIndex > -1 Then
        CurrentValue = ComboBox.ItemData(ComboBox.ListIndex)
    End If
    
    If ClearComboBox Then
        ComboBox.Clear
    Else
        PreloadedItemCount = ComboBox.ListCount
    End If
    
    If Collection.Count > 0 Then
        For Each CollectionItem In Collection
            If TextPropertyName = "" Then
                ComboBox.AddItem CStr(CollectionItem)
            Else
                ComboBox.AddItem CallByName(CollectionItem, TextPropertyName, VbGet)
            End If
            If IDPropertyName <> "" Then
                ComboBox.ItemData(ComboBox.NewIndex) = Val(CallByName(CollectionItem, IDPropertyName, VbGet))
            End If
        Next CollectionItem
    End If
    
    Select Case ItemPosition
        Case cscpNone
            ComboBox.ListIndex = -1
        Case cscpFirst
            If ComboBox.ListCount > 0 Then
                ComboBox.ListIndex = 0
            End If
        Case cscpFirstIfUnique
            If ComboBox.ListCount = 1 Then
                ComboBox.ListIndex = 0
            Else
                ComboBox.ListIndex = -1
            End If
        Case cscpLast
            If ComboBox.ListCount > 0 Then
                ComboBox.ListIndex = ComboBox.ListCount - 1
            End If
        Case cscpItemOrNone, cscpItemOrfirst, cscpItemOrFirstIfUnique, cscpItemOrFirstIfUniqueIgnoringPreloaded, cscpItemOrLast
            ComboBox.ListIndex = GetListIndexByItemData(ComboBox, ItemValue, ItemPosition, PreloadedItemCount)
        Case cscpCurrentOrNone, cscpCurrentOrFirst, cscpCurrentOrFirstIfUnique, cscpCurrentOrFirstIfUniqueIgnoringPreloaded, cscpCurrentOrLast
            ComboBox.ListIndex = GetListIndexByItemData(ComboBox, CurrentValue, ItemPosition, PreloadedItemCount)
    End Select
    
    FillFromCollection = True
End Function

Public Function FillWithHoursAndMinutes(ByRef ComboBox As ComboBox, ByVal StartHour As Boolean, ByVal StartTime As Date, ByVal EndTime As Date, ByVal IntervalMinutes As Byte)
    Dim intMinutos As Integer
    Dim intStart_Hour_Minutes As Integer
    Dim intEnd_Hour_Minutes As Integer
    
    intStart_Hour_Minutes = DateDiff("n", "00:00", StartTime)
    intEnd_Hour_Minutes = DateDiff("n", "00:00", EndTime)
    
    If StartHour Then
        'Es un combo de Horas de Inicio
        For intMinutos = intStart_Hour_Minutes To intEnd_Hour_Minutes Step IntervalMinutes
            ComboBox.AddItem Format(DateAdd("n", intMinutos, "00:00"), "hh:nn")
        Next intMinutos
    Else
        'Es un combo de Horas de Fin
        For intMinutos = IntervalMinutes - 1 To intEnd_Hour_Minutes Step IntervalMinutes
            ComboBox.AddItem Format(DateAdd("n", intMinutos, "00:00"), "hh:nn")
        Next intMinutos
    End If
End Function

Public Function GetListIndexByText(ByRef ComboBox As VB.ComboBox, ByVal TextValue As String, Optional ByVal ItemPosition As csComboPosition = cscpItemOrNone, Optional ByVal ItemLeftCharsToCompare As Byte = 255) As Integer
    Dim ListIndex As Integer
    
    If TextValue <> "" Then
        For ListIndex = 0 To ComboBox.ListCount - 1
            If Left(ComboBox.List(ListIndex), ItemLeftCharsToCompare) = TextValue Then
                GetListIndexByText = ListIndex
                Exit Function
            End If
        Next ListIndex
    End If
    
    Select Case ItemPosition
        Case cscpItemOrNone
            GetListIndexByText = -1
        Case cscpItemOrfirst
            If ComboBox.ListCount > 0 Then
                GetListIndexByText = 0
            Else
                GetListIndexByText = -1
            End If
        Case cscpItemOrLast
            If ComboBox.ListCount > 0 Then
                GetListIndexByText = ComboBox.ListCount - 1
            End If
    End Select
End Function

Public Function GetListIndexByItemData(ByRef ComboBox As VB.ComboBox, ByVal ItemData As Long, Optional ByVal ItemPosition As csComboPosition = cscpItemOrNone, Optional PreloadedItemCount As Long = 0) As Integer
    Dim ListIndexCurrent As Integer
    
    For ListIndexCurrent = 0 To ComboBox.ListCount - 1
        If ItemData = ComboBox.ItemData(ListIndexCurrent) Then
            GetListIndexByItemData = ListIndexCurrent
            Exit Function
        End If
    Next ListIndexCurrent
    
    Select Case ItemPosition
        Case cscpItemOrNone, cscpCurrentOrNone
            GetListIndexByItemData = -1
        Case cscpItemOrfirst, cscpCurrentOrFirst
            If ComboBox.ListCount = 0 Then
                GetListIndexByItemData = -1 'No hay items, por lo tanto devuelvo -1 (ninguno)
            Else
                GetListIndexByItemData = 0  'Hay al menos un item, devuelvo el primero
            End If
        Case cscpItemOrFirstIfUnique, cscpCurrentOrFirstIfUnique
            If ComboBox.ListCount <> 1 Then
                GetListIndexByItemData = -1 'O no hay items, o hay más de uno, devuelvo -1 (ninguno)
            Else
                GetListIndexByItemData = 0  'Hay un sólo item, lo devuelvo.
            End If
        Case cscpItemOrFirstIfUniqueIgnoringPreloaded, cscpCurrentOrFirstIfUniqueIgnoringPreloaded
            If ComboBox.ListCount - PreloadedItemCount <> 1 Then
                GetListIndexByItemData = -1 'O no hay items, o hay más de uno, devuelvo -1 (ninguno)
            Else
                GetListIndexByItemData = 0  'Hay un sólo item, lo devuelvo.
            End If
        Case cscpItemOrLast, cscpCurrentOrLast
            If ComboBox.ListCount > 0 Then
                GetListIndexByItemData = ComboBox.ListCount - 1
            End If
    End Select
End Function

Public Function AndCollection_FillFromSQL(ByRef ComboBox As VB.ComboBox, ByRef CCollection As Collection, ByVal Clear As Boolean, ByVal RecordSource As String, ByVal BoundField As String, ByVal ListField As String, ByVal ErrorEntityName As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Long = 0) As Boolean
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
    
    AndCollection_FillFromSQL = AndCollection_FillFromCommand(ComboBox, CCollection, Clear, cmdData, BoundField, ListField, ErrorEntityName, ItemPosition, ItemValue)
    
    Set cmdData = Nothing
    
    Screen.MousePointer = MousePointerSave
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Modules.Control_Module.AndCollection_FillFromSQL", "Error al leer la lista de " & ErrorEntityName & "."
End Function

Public Function AndCollection_FillFromCommand(ByRef ComboBox As VB.ComboBox, ByRef CCollection As Collection, ByVal Clear As Boolean, ByRef command As ADODB.command, ByVal BoundField As String, ByVal ListField As String, ByVal ErrorEntityName As String, Optional ItemPosition As csComboPosition = cscpNone, Optional ItemValue As Long = 0) As Boolean
    Dim MousePointerSave As Integer
    Dim recData As ADODB.Recordset
    
    If pTrapErrors Then
        On Error GoTo ErrorHandler
    End If
    
    MousePointerSave = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    If Clear Then
        ComboBox.Clear
        Set CCollection = New Collection
    Else
        If CCollection Is Nothing Then
            Set CCollection = New Collection
        End If
    End If
    
    Set recData = New ADODB.Recordset
    recData.Open command, , , adLockReadOnly
    
    Do While Not recData.EOF
        CCollection.Add recData(BoundField).Value, KEY_STRINGER & recData(BoundField).Value
        ComboBox.AddItem recData(ListField).Value
        recData.MoveNext
    Loop
    
    Select Case ItemPosition
        Case cscpNone
            ComboBox.ListIndex = 0
        Case cscpFirst
            If ComboBox.ListCount > 0 Then
                ComboBox.ListIndex = 0
            End If
        Case cscpLast
            If ComboBox.ListCount > 0 Then
                ComboBox.ListIndex = ComboBox.ListCount - 1
            End If
        Case cscpItemOrNone, cscpItemOrfirst, cscpItemOrLast
            ComboBox.ListIndex = CSM_Control_ComboBox.GetListIndexByItemData(ComboBox, ItemValue, ItemPosition)
    End Select
    
    Set recData = Nothing
    Set command = Nothing
    
    Screen.MousePointer = MousePointerSave
    AndCollection_FillFromCommand = True
    Exit Function
    
ErrorHandler:
    ShowErrorMessage "Modules.Control_Module.AndCollection_FillFromCommand", "Error al leer la lista de " & ErrorEntityName & "."
End Function

Public Function AndCollection_GetListIndexByCollectionItem(ByRef ComboBox As VB.ComboBox, ByRef CCollection As Collection, ByVal TextValue As String, Optional ByVal ItemPosition As csComboPosition = cscpItemOrNone) As Long
    Dim CollectionIndex As Long
    
    For CollectionIndex = 1 To CCollection.Count
        If CCollection(CollectionIndex) = TextValue Then
            AndCollection_GetListIndexByCollectionItem = CollectionIndex
            Exit Function
        End If
    Next CollectionIndex
    
    Select Case ItemPosition
        Case cscpItemOrNone
            AndCollection_GetListIndexByCollectionItem = -1
        Case cscpItemOrfirst
            If ComboBox.ListCount > 0 Then
                AndCollection_GetListIndexByCollectionItem = 0
            Else
                AndCollection_GetListIndexByCollectionItem = -1
            End If
        Case cscpItemOrLast
            If ComboBox.ListCount > 0 Then
                AndCollection_GetListIndexByCollectionItem = ComboBox.ListCount - 1
            End If
    End Select
End Function

Public Sub AddItemWithItemData(ByRef ComboBox As VB.ComboBox, ByVal TextItem As String, ByVal ItemData As Long)
    ComboBox.AddItem TextItem
    ComboBox.ItemData(ComboBox.NewIndex) = ItemData
End Sub

Public Function GetSubID(ByRef ComboBox As VB.ComboBox, ByVal SubIDLen As Byte, ByVal SubIDToGet As Byte) As Long
    Dim ComboString As String
    
    If ComboBox.ListIndex > -1 Then
        ComboString = CStr(ComboBox.ItemData(ComboBox.ListIndex))
        
        Select Case SubIDToGet
            Case 1
                GetSubID = Val(Left(ComboString, Len(ComboString) - SubIDLen))
            Case 2
                GetSubID = Val(Right(ComboString, SubIDLen))
            Case Else
        End Select
    End If
End Function

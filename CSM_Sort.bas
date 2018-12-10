Attribute VB_Name = "CSM_Sort"
Option Explicit

'Performs a swap between two values
Private Sub Swap(ByRef Value1, ByRef Value2)
    Dim Temp As Variant
    
    Temp = Value1
    Value1 = Value2
    Value2 = Temp
End Sub

'Splits a subarray into two more subarrays
Private Sub SplitArrayString(ByRef ArrayToSplit() As String, ByVal First As Long, ByVal Last As Long, ByRef LowRet As Long, ByRef HighRet As Long)
    Dim Low As Long
    Dim High As Long
    Dim v As String
    
    v = ArrayToSplit((First + Last) / 2)
    Low = First
    High = Last
    
    Do
        Do While ArrayToSplit(Low) < v
            Low = Low + 1
        Loop
        Do While ArrayToSplit(High) > v
            High = High - 1
        Loop
        If Low <= High Then
            Swap ArrayToSplit(Low), ArrayToSplit(High)
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    LowRet = Low
    HighRet = High
End Sub

Public Sub Sort_ArrayString_Bubble(ByRef ArrayToSort() As String, ByVal First As Long, ByVal Last As Long)
    Dim i As Long
    Dim j As Long
    
    i = First
    Do
        j = Last
        Do
            If ArrayToSort(j) < ArrayToSort(j - 1) Then
                Swap ArrayToSort(j), ArrayToSort(j - 1)
            End If
            j = j - 1
        Loop While j >= i + 1
        i = i + 1
    Loop While i < Last - 1
End Sub

'Performs the QuickSort algorithm
Public Sub Sort_ArrayString_QuickSort(ByRef ArrayToSort() As String, ByVal First As Long, ByVal Last As Long)
    Dim Low As Long
    Dim High As Long
    
    If First < Last Then
        SplitArrayString ArrayToSort, First, Last, Low, High
        If First < High Then
            Sort_ArrayString_QuickSort ArrayToSort, First, High
        End If
        If Low < Last Then
            Sort_ArrayString_QuickSort ArrayToSort, Low, Last
        End If
    End If
End Sub

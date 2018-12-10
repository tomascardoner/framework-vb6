Attribute VB_Name = "CSM_Array"
Option Explicit

Public Sub ClearNumeric(ByRef ArrayToClear As Variant)
    Dim ArrayIndex As Long
    
    For ArrayIndex = LBound(ArrayToClear) To UBound(ArrayToClear)
        ArrayToClear(ArrayIndex) = 0
    Next ArrayIndex
End Sub

Public Sub ClearNumericExcept(ByRef ArrayToClear As Variant, ByVal LowerIndex As Long, ByVal UpperIndex As Long)
    Dim ArrayIndex As Long
    
    For ArrayIndex = LBound(ArrayToClear) To LowerIndex - 1
        ArrayToClear(ArrayIndex) = 0
    Next ArrayIndex
    For ArrayIndex = UpperIndex + 1 To UBound(ArrayToClear)
        ArrayToClear(ArrayIndex) = 0
    Next ArrayIndex
End Sub

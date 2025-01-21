Attribute VB_Name = "std_ArrayFunctions"


Option Explicit

Public Enum ArrOperator
    ArrOperatorAdd = 0
    ArrOperatorSubtract = 1
    ArrOperatorMultiply = 2
    ArrOperatorDivide = 3
    ArrOperatorPower = 4
    ArrOperatorRoot = 5
End Enum

Public Enum ArrCompare
    ArrCompareEqual = 0
    ArrCompareNotEqual = 1
    ArrCompareGreater = 2
    ArrCompareSmaller = 3
    ArrCompareGreaterEqual = 4
    ArrCompareSmallerEqual = 5
End Enum

Public Function MergeArray(Goal() As Variant, Merge() As Variant) As Variant()
    Dim Dimension1    As Long: Dimension1 = ArrayDimension(Goal)
    Dim Dimension2    As Long: Dimension2 = ArrayDimension(Merge)
    Dim NewSizes()    As Long: NewSizes   = ArrayBiggerSize(Goal, Merge)
    Dim GoalSizes()   As Long: GoalSizes  = ArraySizes(Goal)
    Dim MergeSizes()  As Long: MergeSizes = ArraySizes(Merge)
    Dim RestSizes()   As Long
    Dim StartIndex    As Long
    Dim Temp()        As Long
    Dim Temp2()       As Long
    Dim i As Long
    Select Case True
        Case Dimension1 = 0
            MergeArray = Merge
        Case Dimension2 = 0
            MergeArray = Goal
        Case Dimension1 = Dimension2
            StartIndex = GoalSizes(Dimension1 - 1) + 1
            NewSizes(Dimension1 - 1) = StartIndex + MergeSizes(Dimension1 - 1)
            Call CreateArray(MergeArray, NewSizes)

            ReDim Temp(Dimension1 - 1)
            ReDim Temp2(Dimension1 - 1)
            Call AssignArrToDimension(MergeArray, Goal, NewSizes, GoalSizes, Temp, Temp2, 0)

            RestSizes = NewSizes
            Call ArrayOperator(RestSizes, MergeSizes, ArrOperatorSubtract)
            Call AssignArrToDimension(MergeArray, Merge, NewSizes, MergeSizes, RestSizes, Temp, 0)
        Case Else
    End Select
End Function

Public Function SizeDifference(Arr1() As Variant, Arr2() As Variant) As Long()
    Dim Dimension1 As Long: Dimension1 = ArrayDimension(Arr1)
    Dim Dimension2 As Long: Dimension2 = ArrayDimension(Arr2)
    Dim i As Long
    Select Case True
        Case Dimension1 = 0, Dimension2 = 0
        Case Dimension1 <> Dimension2
        Case Else
            ReDim SizeDifference(Dimension1)
            For i = 0 To SizeDifference
                SizeDifference(0) = Ubound(Arr1, i) - Ubound(Arr2, i)
            Next i
    End Select
End Function

Public Function ArrayBiggerSize(Arr1() As Variant, Arr2() As Variant) As Long()
    Dim Dimension1 As Long: Dimension1 = ArrayDimension(Arr1)
    Dim Dimension2 As Long: Dimension2 = ArrayDimension(Arr2)
    Dim i As Long
    Dim Temp() As Long
    Select Case True
        Case Dimension1 = 0, Dimension2 = 0
        Case Dimension1 <> Dimension2
        Case Else
            ReDim Temp(Dimension1 - 1)
            For i = 0 To USize(Temp)
                If USize(Arr1, i + 1) >= USize(Arr2, i + 1) Then 
                    Temp(i) = USize(Arr1, i + 1)
                Else
                    Temp(i) = USize(Arr2, i + 1)
                End If
            Next i
    End Select
    ArrayBiggerSize = Temp
End Function

Public Function ArraySizes(Arr() As Variant) As Long()
    Dim i As Long
    Dim Size As Long
    Dim Temp() As Long
    Size = ArrayDimension(Arr)
    If Size = 0 Then Exit Function
    ReDim Temp(Size - 1)
    For i = 0 To USize(Temp)
        Temp(i) = USize(Arr, i + 1)
    Next i
    ArraySizes = Temp
End Function

Public Sub ArrayOperator(Arr1 As Variant, Arr2 As Variant, Operator As ArrOperator)
    Dim i As Long
    If USize(Arr1) = -1 Then Exit Sub
    For i = 0 To USize(Arr1)
        Select Case Operator
            Case ArrOperator.ArrOperatorAdd      : Arr1(i) = Arr1(i) + Arr2(i)
            Case ArrOperator.ArrOperatorSubtract : Arr1(i) = Arr1(i) - Arr2(i)
            Case ArrOperator.ArrOperatorMultiply : Arr1(i) = Arr1(i) * Arr2(i)
            Case ArrOperator.ArrOperatorDivide   : Arr1(i) = Arr1(i) / Arr2(i)
            Case ArrOperator.ArrOperatorPower    : Arr1(i) = Arr1(i) ^ Arr2(i)
        End Select
    Next i
End Sub

Public Function ArrayCompare(Arr1 As Variant, Arr2 As Variant, Operator As ArrCompare) As Boolean
    Dim i As Long
    For i = 0 To USize(Arr1)
        Select Case Operator
            Case ArrCompare.ArrCompareEqual        : If Arr1(i) <> Arr2(i) Then Exit Function
            Case ArrCompare.ArrCompareNotEqual     : If Arr1(i) =  Arr2(i) Then Exit Function
            Case ArrCompare.ArrCompareGreater      : If Arr1(i) =< Arr2(i) Then Exit Function
            Case ArrCompare.ArrCompareSmaller      : If Arr1(i) >= Arr2(i) Then Exit Function
            Case ArrCompare.ArrCompareGreaterEqual : If Arr1(i) <  Arr2(i) Then Exit Function
            Case ArrCompare.ArrCompareSmallerEqual : If Arr1(i) >  Arr2(i) Then Exit Function
        End Select
    Next i
    ArrayCompare = True
End Function

Public Function CreateArray(Arr() As Variant, Sizes() As Long) As Variant()
    Select Case USize(Sizes)
        Case -01
        Case 00: ReDim Arr(Sizes(0))
        Case 01: ReDim Arr(Sizes(0), Sizes(1))
        Case 02: ReDim Arr(Sizes(0), Sizes(1), Sizes(2))
        Case 03: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3))
        Case 04: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4))
        Case 05: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5))
        Case 06: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6))
        Case 07: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7))
        Case 08: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8))
        Case 09: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9))
        Case 10: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10))
        Case 11: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11))
        Case 12: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12))
        Case 13: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13))
        Case 14: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14))
        Case 15: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15))
        Case 16: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16))
        Case 17: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17))
        Case 18: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18))
        Case 19: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19))
        Case 20: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20))
        Case 21: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21))
        Case 22: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22))
        Case 23: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23))
        Case 24: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23), Sizes(24))
        Case 25: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23), Sizes(24), Sizes(25))
        Case 26: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23), Sizes(24), Sizes(25), Sizes(26))
        Case 27: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23), Sizes(24), Sizes(25), Sizes(26), Sizes(27))
        Case 28: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23), Sizes(24), Sizes(25), Sizes(26), Sizes(27), Sizes(28))
        Case 29: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23), Sizes(24), Sizes(25), Sizes(26), Sizes(27), Sizes(28), Sizes(29))
        Case 30: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23), Sizes(24), Sizes(25), Sizes(26), Sizes(27), Sizes(28), Sizes(29), Sizes(30))
        Case 31: ReDim Arr(Sizes(0), Sizes(1), Sizes(2), Sizes(3), Sizes(4), Sizes(5), Sizes(6), Sizes(7), Sizes(8), Sizes(9), Sizes(10), Sizes(11), Sizes(12), Sizes(13), Sizes(14), Sizes(15), Sizes(16), Sizes(17), Sizes(18), Sizes(19), Sizes(20), Sizes(21), Sizes(22), Sizes(23), Sizes(24), Sizes(25), Sizes(26), Sizes(27), Sizes(28), Sizes(29), Sizes(30), Sizes(31))
        Case Else
            MsgBox("Why the hell do you need more than 32 Dimensions?")
    End Select
End Function

Public Function USize(Arr As Variant, Optional Dimension As Long = 1)
    On Error Resume Next
    USize = -1
    USize = Ubound(Arr, Dimension)
End Function

Public Function ArrayDimension(Arr() As Variant) As Long
    Dim i As Long
    i = 1
    Do Until USize(Arr, i) = -1
        ArrayDimension = i
        i = i + 1
    Loop
End Function

Public Sub AssignArrToDimension(ByRef Arr1() As Variant, ByRef Arr2() As Variant, ByRef Arr1Dim() As Long, ByRef Arr2Dim() As Long, ByRef Arr1CurDim() As Long, ByRef Arr2CurDim() As Long, ByVal CurrentDimension As Long)
    Dim i As Long
    Dim Temp(1) As Long

    Temp(0) = Arr1CurDim(CurrentDimension)
    Temp(1) = Arr2CurDim(CurrentDimension)
    For i = Arr2CurDim(CurrentDimension) To Arr2Dim(CurrentDimension)
        If USize(Arr2Dim) <> CurrentDimension Then Call AssignArrToDimension(Arr1, Arr2, Arr1Dim, Arr2Dim, Arr1CurDim, Arr2CurDim, CurrentDimension + 1)
        Call ArrayLet(Arr1, Arr1CurDim, ArrayGet(Arr2, Arr2CurDim))
        Arr1CurDim(CurrentDimension) = Arr1CurDim(CurrentDimension) + 1
        Arr2CurDim(CurrentDimension) = Arr2CurDim(CurrentDimension) + 1
    Next i
    Arr1CurDim(CurrentDimension) = Temp(0)
    Arr2CurDim(CurrentDimension) = Temp(1)
    
End Sub

Public Sub ArrayLet(Arr() As Variant, Dimensions() As Long, Value As Variant)
    Select Case USize(Dimensions)
        Case -01
        Case 00: Arr(Dimensions(0)) = Value
        Case 01: Arr(Dimensions(0), Dimensions(1)) = Value
        Case 02: Arr(Dimensions(0), Dimensions(1), Dimensions(2)) = Value
        Case 03: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3)) = Value
        Case 04: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4)) = Value
        Case 05: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5)) = Value
        Case 06: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6)) = Value
        Case 07: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7)) = Value
        Case 08: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8)) = Value
        Case 09: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9)) = Value
        Case 10: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10)) = Value
        Case 11: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11)) = Value
        Case 12: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12)) = Value
        Case 13: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13)) = Value
        Case 14: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14)) = Value
        Case 15: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15)) = Value
        Case 16: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16)) = Value
        Case 17: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17)) = Value
        Case 18: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18)) = Value
        Case 19: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19)) = Value
        Case 20: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20)) = Value
        Case 21: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21)) = Value
        Case 22: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22)) = Value
        Case 23: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23)) = Value
        Case 24: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24)) = Value
        Case 25: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25)) = Value
        Case 26: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26)) = Value
        Case 27: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27)) = Value
        Case 28: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27), Dimensions(28)) = Value
        Case 29: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27), Dimensions(28), Dimensions(29)) = Value
        Case 30: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27), Dimensions(28), Dimensions(29), Dimensions(30)) = Value
        Case 31: Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27), Dimensions(28), Dimensions(29), Dimensions(30), Dimensions(31)) = Value
        Case Else
    End Select
End Sub

Public Function ArrayGet(Arr() As Variant, Dimensions() As Long) As Variant
    Select Case USize(Dimensions)
        Case -01
        Case 00: ArrayGet = Arr(Dimensions(0))
        Case 01: ArrayGet = Arr(Dimensions(0), Dimensions(1))
        Case 02: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2))
        Case 03: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3))
        Case 04: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4))
        Case 05: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5))
        Case 06: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6))
        Case 07: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7))
        Case 08: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8))
        Case 09: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9))
        Case 10: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10))
        Case 11: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11))
        Case 12: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12))
        Case 13: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13))
        Case 14: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14))
        Case 15: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15))
        Case 16: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16))
        Case 17: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17))
        Case 18: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18))
        Case 19: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19))
        Case 20: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20))
        Case 21: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21))
        Case 22: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22))
        Case 23: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23))
        Case 24: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24))
        Case 25: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25))
        Case 26: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26))
        Case 27: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27))
        Case 28: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27), Dimensions(28))
        Case 29: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27), Dimensions(28), Dimensions(29))
        Case 30: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27), Dimensions(28), Dimensions(29), Dimensions(30))
        Case 31: ArrayGet = Arr(Dimensions(0), Dimensions(1), Dimensions(2), Dimensions(3), Dimensions(4), Dimensions(5), Dimensions(6), Dimensions(7), Dimensions(8), Dimensions(9), Dimensions(10), Dimensions(11), Dimensions(12), Dimensions(13), Dimensions(14), Dimensions(15), Dimensions(16), Dimensions(17), Dimensions(18), Dimensions(19), Dimensions(20), Dimensions(21), Dimensions(22), Dimensions(23), Dimensions(24), Dimensions(25), Dimensions(26), Dimensions(27), Dimensions(28), Dimensions(29), Dimensions(30), Dimensions(31))
        Case Else
    End Select
End Function

Public Function Flip2D(Arr() As Variant) As Variant()
    Dim Temp() As Variant
    ReDim Temp(USize(Arr, 2), USize(Arr, 1))
    Dim i As Long, j As Long
    For i = 0 To USize(Arr, 1)
        For j = 0 To USize(Arr, 2)
            Temp(j, i) = Arr(i, j)
        Next j
    Next i
    Flip2D = Temp
End Function

Public Sub Sort2D(Arr() As Variant, Optional SortIndex As Long = 0, Optional SortByRow As Boolean = True)
    Dim i As Long, j As Long, k As Long
    Dim Temp As Variant

    If SortByRow Then
        For i = 0 To USize(Arr, 1) - 1
            For j = i To USize(Arr, 1) - i - 1
                If Arr(j, SortIndex) > Arr(j + 1, SortIndex) Then
                    For k = 0 To USize(Arr, 2)
                        Temp = Arr(j, k)
                        Arr(j, k) = Arr(j + 1, k)
                        Arr(j + 1, k) = Temp
                    Next k
                End If
            Next j
        Next i
    Else
        For i = 0 To USize(Arr, 2) - 1
            For j = i To USize(Arr, 2) - i - 1
                If Arr(SortIndex, j) > Arr(SortIndex, j + 1) Then
                    For k = 0 To USize(Arr, 2)
                        Temp = Arr(j, k)
                        Arr(j, k) = Arr(j + 1, k)
                        Arr(j + 1, k) = Temp
                    Next k
                End If
            Next j
        Next i
    End If
End Sub
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

Private Sub Assign(Var1 As Variant, Var2 As Variant)
    If IsObject(Var2) Then
        Set Var1 = Var2
    Else
        Var1 = Var2
    End If
End Sub
' 1D Array
    Public Function ArrayPop(Arr As Variant) As Variant
        Dim Temp As Variant
        Dim i As Long
        Temp = Arr
        ReDim Temp(ArraySize(Temp) - 1)
        For i = 0 To ArraySize(Temp)
            Call Assign(Temp(i), Arr(i))
        Next i
        ArrayPop = Temp
    End Function

    Public Function ArrayPush(Arr As Variant, Value As Variant) As Variant
        Dim Temp As Variant
        Temp = Arr
        ReDim Preserve Temp(ArraySize(Temp) + 1)
        Call Assign(Temp(ArraySize(Temp)), Value)
        ArrayPush = Temp
    End Function

    Public Function ArrayShift(Arr As Variant) As Variant
        Dim Temp As Variant
        Dim i As Long
        If ArraySize(Arr) > 0 Then
            Temp = Arr
            ReDim Temp(ArraySize(Temp) - 1)
            For i = 0 To ArraySize(Temp)
                Call Assign(Temp(i), Arr(i + 1))
            Next i
            ArrayShift = Temp
        End If
    End Function

    Public Function ArrayUnshift(Arr As Variant, Value As Variant) As Variant
        Dim Temp As Variant
        Dim i As Long
        Temp = Arr
        ReDim Temp(ArraySize(Temp) + 1)
        Call Assign(Temp(0), Value)
        For i = 1 To ArraySize(Temp)
            Call Assign(Temp(i), Arr(i))
        Next i
        ArrayUnshift = Temp
    End Function

    Public Function ArrayInsert(Arr As Variant, Value As Variant, Index As Long) As Variant
        Dim Temp As Variant
        Dim i As Long
        Temp = Arr
        ReDim Temp(ArraySize(Temp) + 1)
        For i = 0 To Index - 1
            Call Assign(Temp(i), Arr(i))
        Next i
        Call Assign(Temp(Index), Value)
        For i = Index To ArraySize(Temp)
            Call Assign(Temp(i), Arr(i - 1))
        Next i
        ArrayInsert = Temp
    End Function

    Public Function ArrayRemove(Arr As Variant, Index As Long) As Variant
        Dim Temp As Variant
        Dim i As Long
        Temp = Arr
        ReDim Temp(ArraySize(Temp) - 1)
        For i = 0 To Index - 1
            Call Assign(Temp(i), Arr(i))
        Next i
        For i = Index To ArraySize(Temp)
            Call Assign(Temp(i), Arr(i + 1))
        Next i
        ArrayRemove = Temp
    End Function

    Public Function ArraySplice(Arr As Variant, StartIndex As Long, Optional EndIndex As Long = -1) As Variant
        Dim Temp As Variant
        Dim i As Long
        Dim Index As Long
        If EndIndex = -1 Then EndIndex = ArraySize(Arr)
        Temp = Arr
        ReDim Temp(ArraySize(Arr) - (EndIndex - StartIndex + 1))
        For i = 0 To ArraySize(Arr)
                If StartIndex > i Or EndIndex < i Then
                    Call Assign(Temp(Index), Arr(i))
                Index = Index + 1
            End If
        Next i
        ArraySplice = Temp
    End Function

    Public Function ArrayJoin(Arr As Variant, Text As String) As String
        Dim i As Long
        Dim Message As String
        For i = 0 To ArraySize(Temp)
            Message = Message & CStr(Arr(i)) & Text
        Next i
        Message = Mid(Message, 1, Len(Message) - Len(Text))
        ArrayJoin = Message
    End Function

    Public Function ArrayReverse(Arr As Variant) As Variant
        Dim Temp As Variant
        Dim TempVar As Variant
        Dim i As Long
        Temp = Arr
        For i = 0 To ArraySize(Temp)
            Call Assign(TempVar, Temp(i))
            Call Assign(Temp(i), Temp(ArraySize(Temp) - 1))
            Call Assign(Temp(ArraySize(Temp) - 1), TempVar)
        Next i
        ArrayReverse = Temp
    End Function

    Public Function ArrayFill(Arr As Variant, Value As Variant) As Variant
        Dim Temp As Variant
        Dim i As Long
        Temp = Arr
        For i = 0 To ArraySize(Temp)
            Call Assign(Temp(i), Value)
        Next i
        ArrayFill = Temp
    End Function

    Public Function ArrayIndex(Arr As Variant, Value As Variant) As Long
        Dim Temp As Variant
        Dim i As Long
        Temp = Arr
        ArrayIndex = -1
        For i = 0 To ArraySize(Temp)
            If Temp(i) = Value Then
                ArrayIndex = i
                Exit Function
            End If
        Next i
    End Function

    Public Function ArrayLastIndex(Arr As Variant, Value As Variant) As Long
        Dim Temp As Variant
        Dim i As Long
        Temp = Arr
        ArrayIndex = -1
        For i = ArraySize(Temp) To 0 Step-1
            If Temp(i) = Value Then
                ArrayIndex = i
                Exit Function
            End If
        Next i
    End Function

    Public Function ArrayIncludes(Arr As Variant, Value As Variant) As Boolean
        ArrayIncludes = ArrayIndex(Arr, Value) > -1
    End Function

    Public Function ArrayInsertArray(Arr As Variant, Insert As Variant, Position As Long) As Variant
        Dim Temp As Variant
        Dim i As Long
        Dim NewSize As Long
        Dim CurrentIndex As Long

        Temp = Arr
        NewSize = (ArraySize(Arr) + 1 + ArraySize(Insert) + 1) - 1
        ReDim Temp(NewSize)
        For i = 0 To Position
            Call Assign(Temp(i), Arr(i))
        Next i
        CurrentIndex = i
        For i = 0 To ArraySize(Insert)
            Call Assign(Temp(CurrentIndex + i), Insert(i))
        Next i
        CurrentIndex = CurrentIndex + i
        For i = Position To ArraySize(Insert)
            Call Assign(Temp(CurrentIndex + i - Position), Arr(i))
        Next i
        ArrayInsertArray = Temp
    End Function

    Public Function ArrayInsertEach(Arr As Variant, Value As Variant) As Variant
        Dim Temp As Variant
        Dim i As Long
        Dim Index As Long
        Temp = Arr
        ReDim Temp(((ArraySize(Temp) + 1) * 2) - 2)
        For i = 0 To ArraySize(Temp) Step +2
            Call Assign(Temp(i), Arr(Index))
            If i + 1 > ArraySize(Temp) Then Exit For
            Call Assign(Temp(i + 1), Value)
            Index = Index + 1
        Next i
        ArrayInsertEach = Temp
    End Function

    Public Function ArrayConvert(Arr As Variant, ConvertTo As VbVarType) As Variant
        Dim i As Long
        Dim ReturnArray As Variant

        On Error GoTo Error
        ReDim ReturnArray(ArraySize(Arr))
        For i = 0 To ArraySize(Arr)
            Select Case ConvertTo
            Case vbEmpty      : ReturnArray(i) = Empty
            Case vbNull       : ReturnArray(i) = Null
            Case vbInteger    : ReturnArray(i) = Cint(Arr(i))
            Case vbLong       : ReturnArray(i) = CLng(Arr(i))
            Case vbSingle     : ReturnArray(i) = CSng(Arr(i))
            Case vbDouble     : ReturnArray(i) = CDbl(Arr(i))
            Case vbCurrency   : ReturnArray(i) = CCur(Arr(i))
            Case vbDate       : ReturnArray(i) = CDate(Arr(i))
            Case vbString     : ReturnArray(i) = CStr(Arr(i))
            Case vbBoolean    : ReturnArray(i) = CBool(Arr(i))
            Case vbVariant    : ReturnArray(i) = CVar(Arr(i))
            Case vbDecimal    : ReturnArray(i) = CDec(Arr(i))
            Case vbByte       : ReturnArray(i) = CByte(Arr(i))
            Case vbLongLong   : ReturnArray(i) = CLngLng(Arr(i))
            End Select
        Next i
        ArrayConvert = ReturnArray
        Exit Function

        Error:
        Dim Temp As Variant
        ArrayConvert = Temp
    End Function

    Public Function ArrayConvertString(Arr As Variant) As String()
        Dim i As Long
        Dim ReturnArray() As String
        On Error GoTo Error
        ReDim ReturnArray(ArraySize(Arr))
        For i = 0 To ArraySize(Arr)
            ReturnArray(i) = CStr(Arr(i))
        Next i
        ArrayConvertString = ReturnArray
        Error:
    End Function

    Public Function ArrayConvertVariant(Arr As Variant) As Variant()
        Dim i As Long
        Dim ReturnArray() As Variant
        On Error GoTo Error
        ReDim ReturnArray(ArraySize(Arr))
        For i = 0 To ArraySize(Arr)
            ReturnArray(i) = CVar(Arr(i))
        Next i
        ArrayConvertVariant = ReturnArray
        Error:
    End Function

    Public Function ArrayMerge(Goal As Variant, Merge As Variant) As Variant
        Dim NewSize As Long
        Dim Index As Long
        Dim i As Long
        Dim ReturnArray As Variant
        Index = ArraySize(Goal) + 1
        NewSize = Index + (ArraySize(Merge) + 1) - 1
        If NewSize > -1 Then
            ReturnArray = Goal
            ReDim Preserve ReturnArray(NewSize)
            For i = 0 To ArraySize(Merge)
                ReturnArray(Index + i) = Merge(i)
            Next i
            ArrayMerge = ReturnArray
        End If
    End Function
'

Public Function SizeDifference(Arr1 As Variant, Arr2 As Variant) As Long()
    Dim Dimension1 As Long: Dimension1 = ArrayDimension(Arr1)
    Dim Dimension2 As Long: Dimension2 = ArrayDimension(Arr2)
    Dim i As Long
    Select Case True
        Case Dimension1 = 0, Dimension2 = 0
        Case Dimension1 <> Dimension2
        Case Else
            ReDim SizeDifference(Dimension1)
            For i = 0 To SizeDifference
                SizeDifference(0) = ArraySize(Arr1, i) - ArraySize(Arr2, i)
            Next i
    End Select
End Function

Public Function ArrayBiggerSize(Arr1 As Variant, Arr2 As Variant) As Long()
    Dim Dimension1 As Long: Dimension1 = ArrayDimension(Arr1)
    Dim Dimension2 As Long: Dimension2 = ArrayDimension(Arr2)
    Dim i As Long
    Dim Temp() As Long
    Select Case True
        Case Dimension1 = 0, Dimension2 = 0
        Case Dimension1 <> Dimension2
        Case Else
            ReDim Temp(Dimension1 - 1)
            For i = 0 To ArraySize(Temp)
                If ArraySize(Arr1, i + 1) >= ArraySize(Arr2, i + 1) Then 
                    Temp(i) = ArraySize(Arr1, i + 1)
                Else
                    Temp(i) = ArraySize(Arr2, i + 1)
                End If
            Next i
    End Select
    ArrayBiggerSize = Temp
End Function

Public Function ArraySizes(Arr As Variant) As Long()
    Dim i As Long
    Dim Size As Long
    Dim Temp() As Long
    Size = ArrayDimension(Arr)
    If Size = 0 Then Exit Function
    ReDim Temp(Size - 1)
    For i = 0 To ArraySize(Temp)
        Temp(i) = ArraySize(Arr, i + 1)
    Next i
    ArraySizes = Temp
End Function

Public Sub ArrayOperator(Arr1 As Variant, Arr2 As Variant, Operator As ArrOperator)
    Dim i As Long
    If ArraySize(Arr1) = -1 Then Exit Sub
    For i = 0 To ArraySize(Arr1)
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
    For i = 0 To ArraySize(Arr1)
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

Public Function CreateArray(Arr As Variant, Sizes() As Long) As Variant()
    Select Case ArraySize(Sizes)
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

Public Function ArraySize(Arr As Variant, Optional Dimension As Long = 1)
    On Error Resume Next
    ArraySize = -1
    ArraySize = Ubound(Arr, Dimension)
End Function

Public Function ArrayDimension(Arr As Variant) As Long
    Dim i As Long
    i = 1
    Do Until ArraySize(Arr, i) = -1
        ArrayDimension = i
        i = i + 1
    Loop
End Function

Public Sub AssignArrToDimension(ByRef Arr1 As Variant, ByRef Arr2 As Variant, ByRef Arr1Dim() As Long, ByRef Arr2Dim() As Long, ByRef Arr1CurDim() As Long, ByRef Arr2CurDim() As Long, ByVal CurrentDimension As Long)
    Dim i As Long
    Dim Temp(1) As Long

    Temp(0) = Arr1CurDim(CurrentDimension)
    Temp(1) = Arr2CurDim(CurrentDimension)
    For i = Arr2CurDim(CurrentDimension) To Arr2Dim(CurrentDimension)
        If ArraySize(Arr2Dim) <> CurrentDimension Then Call AssignArrToDimension(Arr1, Arr2, Arr1Dim, Arr2Dim, Arr1CurDim, Arr2CurDim, CurrentDimension + 1)
        Call ArrayLet(Arr1, Arr1CurDim, ArrayGet(Arr2, Arr2CurDim))
        Arr1CurDim(CurrentDimension) = Arr1CurDim(CurrentDimension) + 1
        Arr2CurDim(CurrentDimension) = Arr2CurDim(CurrentDimension) + 1
    Next i
    Arr1CurDim(CurrentDimension) = Temp(0)
    Arr2CurDim(CurrentDimension) = Temp(1)
    
End Sub

Public Sub ArrayLet(Arr As Variant, Dimensions() As Long, Value As Variant)
    Select Case ArraySize(Dimensions)
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

Public Function ArrayGet(Arr As Variant, Dimensions() As Long) As Variant
    Select Case ArraySize(Dimensions)
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

Public Function Flip2D(Arr As Variant) As Variant()
    Dim Temp() As Variant
    ReDim Temp(ArraySize(Arr, 2), ArraySize(Arr, 1))
    Dim i As Long, j As Long
    For i = 0 To ArraySize(Arr, 1)
        For j = 0 To ArraySize(Arr, 2)
            Temp(j, i) = Arr(i, j)
        Next j
    Next i
    Flip2D = Temp
End Function

Public Sub Sort2D(Arr As Variant, Optional SortIndex As Long = 0, Optional SortByRow As Boolean = True)
    Dim i As Long, j As Long, k As Long
    Dim Temp As Variant

    If SortByRow Then
        For i = 0 To ArraySize(Arr, 1) - 1
            For j = i To ArraySize(Arr, 1) - i - 1
                If Arr(j, SortIndex) > Arr(j + 1, SortIndex) Then
                    For k = 0 To ArraySize(Arr, 2)
                        Temp = Arr(j, k)
                        Arr(j, k) = Arr(j + 1, k)
                        Arr(j + 1, k) = Temp
                    Next k
                End If
            Next j
        Next i
    Else
        For i = 0 To ArraySize(Arr, 2) - 1
            For j = i To ArraySize(Arr, 2) - i - 1
                If Arr(SortIndex, j) > Arr(SortIndex, j + 1) Then
                    For k = 0 To ArraySize(Arr, 2)
                        Temp = Arr(j, k)
                        Arr(j, k) = Arr(j + 1, k)
                        Arr(j + 1, k) = Temp
                    Next k
                End If
            Next j
        Next i
    End If
End Sub
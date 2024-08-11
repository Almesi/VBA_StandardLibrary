Attribute VB_Name = "AL_Sort"

Public Sub AL_Sort_Bubble(ByRef Array1() As Integer)

    Dim Array1Length As Integer
    Array1Length = UBound(Array1) - LBound(Array1)

    For i = 0 To Array1Length
        For j = i + 1 To Array1Length
            If Array1(i) > Array1(j) Then
                AL_SwapInt Array1(i), Array1(j)
            End If
        Next j
    Next i

End Sub

Public Sub AL_Sort_Merge(ByRef Array1() As Integer, LowestIndex As Integer, HighestIndex As Integer)

    ' Check if there are at least two elements in the array to sort
    If LowestIndex < HighestIndex Then
        Dim ArrayMid As Integer
        ArrayMid = Int((LowestIndex + HighestIndex) / 2)

        ' Recursively sort the left and right halves of the array
        AL_SortMergeInt Array1, LowestIndex, ArrayMid
        AL_SortMergeInt Array1, ArrayMid + 1, HighestIndex

        ' Merge the sorted halves of the array
        AL_MergeInt Array1, LowestIndex, ArrayMid, HighestIndex
    End If
End Sub
    
Public Sub AL_Sort_Merge(ByRef Array1() As Integer, LowestIndex As Integer, ArrayMid As Integer, HighestIndex As Integer)
    
    ' Calculate the sizes of the left and right subarrays
    Dim LeftSize As Integer
    Dim RightSize As Integer
    LeftSize = (ArrayMid - LowestIndex + 1)
    RightSize = (HighestIndex - ArrayMid)

    ' Declare arrays to hold the left and right subarrays
    Dim LeftArray() As Integer
    Dim RightArray() As Integer
    ReDim LeftArray(LeftSize - 1)
    ReDim RightArray(RightSize - 1)

    ' Copy data from the original array to the left and right subarrays
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = 0 To LeftSize - 1
        LeftArray(i) = Array1(LowestIndex + i)
    Next i
    For j = 0 To RightSize - 1
        RightArray(j) = Array1(ArrayMid + j + 1)
    Next j

    ' Merge the left and right subarrays back into the original array
    i = 0
    j = 0
    k = LowestIndex
    Do While i < LeftSize And j < RightSize
        If LeftArray(i) <= RightArray(j) Then
            Array1(k) = LeftArray(i)
            i = i + 1
        Else
            Array1(k) = RightArray(j)
            j = j + 1
        End If
        k = k + 1
    Loop

    ' Copy any remaining elements from the left subarray
    Do While i < LeftSize
        Array1(k) = LeftArray(i)
        i = i + 1
        k = k + 1
    Loop

    ' Copy any remaining elements from the right subarray
    Do While j < RightSize
        Array1(k) = RightArray(j)
        j = j + 1
        k = k + 1
    Loop
End Sub

Public Sub AL_Sort_Quick(ByRef Array1() As Integer, Low As Integer, High As Integer)

    If Low < High Then
        Dim Pivot As Integer

        Pivot = AL_Partition(Array1, Low, High)
        AL_SortQuickInt Array1, Low, Pivot - 1
        AL_SortQuickInt Array1, Pivot + 1, High
    End If

End Sub
    
Function AL_Partition(ByRef Array1() As Integer, Low As Integer, High As Integer) As Integer

    Dim Pivot As Integer
    Dim i As Integer
    Dim j As Integer
    
    Pivot = Array1(High)
    i = Low - 1
    For j = Low To High - 1
        If Array1(j) =< Pivot Then
            i = i + 1
            Range("A1").Offset(0, i).Select
            AL_SwapInt Array1(i), Array1(j)
        End If
    Next j
    AL_SwapInt Array1(i + 1), Array1(High)
    AL_Partition = i + 1

End Function
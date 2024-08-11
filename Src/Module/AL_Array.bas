Attribute VB_Name = "AL_Array"

' Delete a certain Element in the Array and switch all higher numbers one down
Public Sub AL_Array_Delete(ByRef Arr() As Variant, DeleteIndex As Long)

    Dim TempArr() As Variant
    Dim ArrLength As Long
    Dim j As Long
    
    ArrLength = UBound(Arr)
    ReDim TempArr(ArrLength)
    j = 0
    TempArr = Arr
    ReDim Arr(ArrLength - 1)
    If ArrLength = DeleteIndex Then
        For i = 0 To ArrLength - 1
            Arr(i) = TempArr(i)
        Next
        Else
            For i = 0 To ArrLength
                If i = DeleteIndex Then
                    i = i + 1
                End If
                Arr(j) = TempArr(i)
                j = j + 1
            Next
    End If

End Sub

' Search for a Value in an Array and return true, if it is found 
Function AL_Array_InArray(ByRef Arr() As Variant, SearchValue As Variant) As Boolean

    For i = 0 To UBound(Arr)
        If Arr(i) = SearchValue Then
            AL_Array_1InArrVariant = True
            Exit Function
        End If
    Next

End Function

' Add second Array on top of first array and return the whole array
Function AL_Array_Merge(ByRef Arr() As Variant, ByRef Arr2() As Variant) As Variant()

    Dim ArrLength As Long
    Dim Arr2Length As Long
    Dim TempArr() As Variant

    ArrLength = UBound(Arr)
    Arr2Length = UBound(Arr2)
    ReDim TempArr(ArrLength + Arr2Length + 1)
    For i = 0 To ArrLength
        TempArr(i) = Arr(i)
    Next
    For j = 0 To Arr2Length - 1
        TempArr(i + j + 1) = Arr2(j)
    Next
    AL_Array_Merge = TempArr

End Function

' Remove the uppermost value
Public Sub AL_Array_Pop(ByRef Arr() As Variant)

    Dim ArrLength As Long 
    Dim NewArrLength As Long
    Dim TempArr() As Variant

    ArrLength = UBound(Arr)
    NewArrLength = ArrLength - 1
    ReDim TempArr(NewArrLength)
    For i = 0 To NewArrLength
        TempArr(i) = Arr(i)
    Next
    ReDim Arr(NewArrLength)
    Arr = TempArr

End Sub

' Add new Element on top of array
Public Sub AL_Array_Push(ByRef Arr() As Variant, PushObject As Variant)

    Dim ArrLength As Long

    ArrLength = UBound(Arr) + 1
    ReDim Preserve Arr(ArrLength)
    Arr(ArrLength) = PushObject

End Sub

' Search a value in an array and return Index where it is found. If -1 then nothing was found
Function AL_Array_Search(ByRef Arr() As Variant, SearchValue As Variant) As Long

    For i = 0 To UBound(Arr)
        If Arr(i) = SearchValue Then
            AL_Array_1SearchVariant = i
            Exit Function
        End If
    Next
    AL_Array_Search = -1

End Function

' Delete first element of array
Public Sub AL_Array_Shift(ByRef Arr() As Variant)

    Dim ArrLength As Long
    Dim NewArrLength As Long
    Dim TempArr() As Variant

    ArrLength = UBound(Arr)
    NewArrLength = ArrLength - 1
    ReDim TempArr(NewArrLength)
    For i = 0 To NewArrLength
        TempArr(i) = Arr(i + 1)
    Next
    ReDim Arr(NewArrLength)
    Arr = TempArr

End Sub

' Insert value at Index
Public Sub AL_Array_Splice(ByRef Arr() As Variant, SpliceValue As Variant, Index As Long)

    Dim ArrLength As Long

    ArrLength = UBound(Arr) + 1
    ReDim Preserve Arr(ArrLength)
    For i = ArrLength To Index + 1 Step-1
        Arr(i) = Arr(i - 1)
    Next
    Arr(Splice) = SpliceValue

End Sub

' Add value new element of beginning of array
Public Sub AL_Array_Unshift(ByRef Arr() As Variant, UnshiftObject As Variant)
    
    Dim ArrLength As Long

    ArrLength = UBound(Arr) + 1
    ReDim Preserve Arr(ArrLength)
    For i = ArrLength To 1 Step -1
        Arr(i) = Arr(i - 1)
    Next i
    Arr(0) = UnshiftObject

End Sub
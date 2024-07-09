Attribute VB_Name = "AL_Math"

Private Function AL_Math_Difference(Number1 As Variant, Number2 As Variant)

    If Number1 < 0 And Number2 < 0 Then
            Number1 = Abs(Number1)
            Number2 = Abs(Number2)
            AL_Math_Difference = Abs(Number1 - Number2)
        ElseIf Number1 < 0 Then
            AL_Math_Difference = Abs(Number2 - Number1)
        Else
            AL_Math_Difference = Abs(Number1 - Number2)
    End If

End Function

Public Sub AL_RandomArray_Get(ByRef Array1() As Variant, Optional Seed As LongLong = Empty, Optional UpperLimit As Variant = Empty, Optional LowerLimit As Variant = Empty, Optional Unique As Boolean = False) As Boolean

    Dim Temp As Variant

    Randomize Seed
    If UpperLimit = Empty Then
        UpperLimit = 18446744073709551615
    End If
    If LowerLimit = Empty Then
        LowerLimit = -18446744073709551615
    End If

    ' Check if it should only contain unique Values and also check if there is enough room in the Array to save all unique Values inbetween the limits
    If Unique = True Then
                If (UpperLimit - LowerLimit) > UBound(Array1) Then
                    For i = 0 To UBound(Array1)
                        Temp = Int((UpperLimit - LowerLimit + 1) * Rnd + LowerLimit)
                        If AL_Array_InArray_Variant(Array1, Temp) = True Then
                                i = i - 1
                            Else
                                Array1(i) = Temp
                        End If
                    Next
                    AL_RandomArray_Get = True
            End If
        Else
            For i = 0 To UBound(Array1)
                Array1(i) = Int((UpperLimit - LowerLimit + 1) * Rnd + LowerLimit)
            Next
            AL_RandomArray_Get = True
    End If

End Sub

Public Function AL_RandomValue_Get(Optional Seed As LongLong = Empty, Optional UpperLimit As Variant = Empty, Optional LowerLimit As Variant = Empty) As Variant

    Randomize Seed
    If UpperLimit = Empty Then
        UpperLimit = 18446744073709551615
    End If
    If LowerLimit = Empty Then
        LowerLimit = -18446744073709551615
    End If
    AL_RandomValue_Get = Int((UpperLimit - LowerLimit + 1) * Rnd + LowerLimit)

End Function
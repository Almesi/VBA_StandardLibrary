Attribute VB_Name = "std_MathFunctions"

Option Explicit

Public Function Sum(ByRef Values As Variant) As Double
    Dim i As Long
    Dim Result As Double

    For i = LBound(Values) To UBound(Values)
        Result = Result + Values(i)
    Next i

    Sum = Result
End Function

Public Function Average(ByRef Values As Variant) As Double
    Average = MathSum(Values) / (UBound(Values) - LBound(Values) + 1)
End Function

Public Function Count(ByRef Values As Variant) As Long
    Count = UBound(Values) - LBound(Values) + 1
End Function

Public Function Min(ByRef Values As Variant) As Double
    Dim i As Long
    Dim Result As Double

    Result = Values(LBound(Values))

    For i = LBound(Values) + 1 To UBound(Values)
        If Values(i) < Result Then Result = Values(i)
    Next i

    Min = Result
End Function

Public Function Max(ByRef Values As Variant) As Double
    Dim i As Long
    Dim Result As Double

    Result = Values(LBound(Values))

    For i = LBound(Values) + 1 To UBound(Values)
        If Values(i) > Result Then Result = Values(i)
    Next i

    Max = Result
End Function

Public Function Product(ByRef Values As Variant) As Double
    Dim i As Long
    Dim Result As Double

    Result = 1

    For i = LBound(Values) To UBound(Values)
        Result = Result * Values(i)
    Next i

    Product = Result
End Function

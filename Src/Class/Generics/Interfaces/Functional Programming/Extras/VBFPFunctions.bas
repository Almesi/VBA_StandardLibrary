Attribute VB_Name = "VBFPFunctions"

Option Explicit

Public Function Map(ByVal Elements As IFunction, ByVal Func As IFunction) As IFunction
    Dim Source() As IFunction : Source = Elements.Evaluate()
    Dim Size     As Long      : Size   = UBound(Source)
    Dim Result() As IFunction : ReDim Result(Size)

    Dim i As Long
    For i = 0 To Size
        Set Result(i) = Func.Run(Source(i))
    Next i

    Set Map = VBFPValue(Result)
End Function

Public Function Filter(ByVal Elements As IFunction, ByVal Predicate As IFunction) As IFunction
    Dim Source() As IFunction : Source = Elements.Evaluate()
    Dim Result() As IFunction

    Dim Count As Long
    Dim i     As Long
    For i = 0 To UBound(Source)

        If CBool(Predicate.Run(Source(i)).Evaluate()) Then
            ReDim Preserve Result(Count)
            Set Result(Count) = Source(i)
            Count = Count + 1
        End If
    Next i
    Set Filter = VBFPValue(Result)
End Function

Public Function Fold(ByVal Elements As IFunction, ByVal Func As IFunction, ByVal Initial As IFunction) As IFunction
    Dim Source() As IFunction :     Source = Elements.Evaluate()
    Dim Result   As IFunction : Set Result = Initial

    Dim i As Long
    For i = 0 To UBound(Source)
        Set Result = Func.Run(Result, Source(i))
    Next i
    Set Fold = Result
End Function

Public Function Head(ByVal Elements As IFunction) As IFunction
    Dim Source() As IFunction
    Source = Elements.Evaluate()

    Set Head = Source(0)
End Function

Public Function Tail(ByVal Elements As IFunction) As IFunction
    Dim Source() As IFunction
    Source = Elements.Evaluate()

    Dim Result() As IFunction
    Dim i As Long
    Dim j As Long

    If UBound(Source) <= 0 Then
        Set Tail = VBFPValue(Result)
        Exit Function
    End If

    ReDim Result(UBound(Source) - 1)

    j = 0

    For i = 0 + 1 To UBound(Source)
        Set Result(j) = Source(i)
        j = j + 1
    Next i

    Set Tail = VBFPValue(Result)
End Function

' Only pass in IFunction for Functions
Public Function Pipe(ByVal Value As IFunction, ParamArray Functions() As Variant) As IFunction
    Dim Result As IFunction
    Set Result = Value

    Dim i      As Long
    For i = 0 To UBound(Functions)
        Set Result = Functions(i).Run(Result)
    Next i
    Set Pipe = Result
End Function
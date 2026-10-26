Attribute VB_Name = "VBFPHelpers"

Option Explicit

Public Function VBFPValueF(ByVal Value As Variant) As IFunction
    Set VBFPValueF = VBFPValue(Value)
End Function
Public Function VBFPValue(ByVal Value As Variant) As ValueFunction
    Set VBFPValue = ValueFunction.Create(Value)
End Function

Public Function VBFPNamedF(ByVal Name As String) As IFunction
    Set VBFPNamedF = VBFPNamed(Name)
End Function
Public Function VBFPNamed(ByVal Name As String) As NamedFunction
    Set VBFPNamed = NamedFunction.Create(Name)
End Function

Public Function VBFPBoundF(ByVal Func As IFunction, ParamArray Arguments() As Variant) As IFunction
    Dim Temp() As Variant
    Temp = Arguments

    Set VBFPBoundF = BoundFunction.CreateArr(Func, Temp)
End Function
Public Function VBFPBound(ByVal Func As IFunction, ParamArray Arguments() As Variant) As BoundFunction
    Dim Temp() As Variant
    Dim Passed() As IFunction

    Temp = Arguments
    Passed = ArgumentsToFunctions(Temp)

    Set VBFPBound = BoundFunction.CreateArr(Func, Passed)
End Function

Public Function VBFPComposedF(ByVal Outer As IFunction, ByVal Inner As IFunction) As IFunction
    Set VBFPComposedF = ComposedFunction.Create(Outer, Inner)
End Function
Public Function VBFPComposed(ByVal Outer As IFunction, ByVal Inner As IFunction) As BoundFunction
    Set VBFPComposedF = ComposedFunction.Create(Outer, Inner)
End Function

Public Function VBFPCurriedF(ByVal Func As IFunction, ParamArray Arguments() As Variant) As IFunction
    Dim Temp() As Variant
    Temp = Arguments
    Set VBFPCurriedF = CurriedFunction.CreateArr(Func, Temp)
End Function
Public Function VBFPCurried(ByVal Func As IFunction, ParamArray Arguments() As Variant) As BoundFunction
    Dim Temp() As Variant
    Temp = Arguments
    Set VBFPCurried = CurriedFunction.CreateArr(Func, Temp)
End Function

Public Function VBFPMemoizedF(ByVal Func As IFunction) As IFunction
    Set VBFPMemoizedF = VBFPMemoized(Func)
End Function
Public Function VBFPMemoized(ByVal Func As IFunction) As MemoizedFunction
    Set VBFPMemoized = MemoizedFunction.Create(Func)
End Function


Public Sub VBFPPrint(ByVal Value As IFunction)
    Dim Evaluated As Variant
    Dim Element   As Variant

    Evaluated = Value.Evaluate()

    If IsObject(Evaluated) Then
        If TypeOf Evaluated Is IFunction Then
            Call VBFPPrint(Evaluated)
            Exit Sub
        End If
    End If

    If IsArray(Evaluated) Then
        For Each Element In Evaluated
            If TypeOf Element Is IFunction Then
                Call VBFPPrint(Element)
            Else
                Debug.Print Element
            End If
        Next Element
    Else
        Debug.Print Evaluated
    End If
End Sub

Public Sub VBFPPrintExtra(ByVal ExtraText As String, ByVal Value As IFunction)
    Debug.Print ExtraText
    Call VBFPPrint(Value)
End Sub

Public Sub VBFPAssign(ByRef Goal As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Goal = Value
    Else
        Let Goal = Value
    End If
End Sub

Public Function ArgumentsToFunctions(ByRef Arguments() As Variant) As IFunction()
    Dim Result() As IFunction
    Dim Count As Long
    Dim i As Long

    Count = USize(Arguments)

    If Count < 0 Then
        ArgumentsToFunctions = Result
        Exit Function
    End If

    ReDim Result(Count)

    For i = 0 To Count
        Set Result(i) = Arguments(i)
    Next i

    ArgumentsToFunctions = Result
End Function

Public Function FuncArgumentsToArguments(ByRef Arguments() As Variant, ByVal ExpectIFunction As Boolean, ByVal ArgIsIFunction As Boolean) As Variant()
    Dim Result() As Variant
    Dim Count As Long
    Dim i As Long

    Count = USize(Arguments)
    If Count < 0 Then
        FuncArgumentsToArguments = Result
        Exit Function
    End If

    ReDim Result(Count)

    If ExpectIFunction Then
        For i = 0 To Count
            Dim Temp As IFunction
            Set Temp = Arguments(i)
            If ArgIsIFunction Then
                Call VBFPAssign(Result(i), Temp)
            Else
                Call VBFPAssign(Result(i), Temp.Evaluate())
            End If
        Next i
    Else
        For i = 0 To Count
            Call VBFPAssign(Result(i), Arguments(i))
        Next i
    End If

    FuncArgumentsToArguments = Result
End Function
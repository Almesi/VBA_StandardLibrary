Attribute VB_Name = "VBFPExampleFunctions"

Option Explicit


' ============================================================
' Arithmetic
' ============================================================

Public Function NormalAdd(ByVal A As Long, ByVal B As Long) As Long
    NormalAdd = A + B
End Function

Public Function Add(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Add = VBFPValue(Num1.Evaluate() + Num2.Evaluate())
End Function

Public Function Subtract(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Subtract = VBFPValue(Num1.Evaluate() - Num2.Evaluate())
End Function

Public Function Multiply(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Multiply = VBFPValue(Num1.Evaluate() * Num2.Evaluate())
End Function

Public Function Divide(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Divide = VBFPValue(Num1.Evaluate() / Num2.Evaluate())
End Function

Public Function Modulo(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Modulo = VBFPValue(Num1.Evaluate() Mod Num2.Evaluate())
End Function

Public Function Negate(ByVal Num As IFunction) As IFunction
    Set Negate = VBFPValue(-Num.Evaluate())
End Function

Public Function Absolute(ByVal Num As IFunction) As IFunction
    Set Absolute = VBFPValue(Abs(Num.Evaluate()))
End Function

Public Function Maximum(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Maximum = VBFPValue(Max(Num1.Evaluate(), Num2.Evaluate()))
End Function

Public Function Minimum(ByVal Num1 As IFunction, ByVal Num2 As IFunction) As IFunction
    Set Minimum = VBFPValue(Min(Num1.Evaluate(), Num2.Evaluate()))
End Function

Public Function Factorial(ByVal Count As IFunction) As IFunction
    Dim N As Long
    N = Count.Evaluate()

    If N <= 1 Then
        Set Factorial = VBFPValue(1)
        Exit Function
    End If

    Dim Previous As IFunction
    Dim Result As Long

    Set Previous  = Factorial(VBFPValue(N - 1))
    Result        = N * Previous.Evaluate()
    Set Factorial = VBFPValue(Result)
End Function


' ============================================================
' Comparison
' ============================================================

Public Function Equal(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set Equal = VBFPValue(Value1.Evaluate() = Value2.Evaluate())
End Function

Public Function NotEqual(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set NotEqual = VBFPValue(Value1.Evaluate() <> Value2.Evaluate())
End Function

Public Function GreaterThan(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set GreaterThan = VBFPValue(Value1.Evaluate() > Value2.Evaluate())
End Function

Public Function LessThan(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set LessThan = VBFPValue(Value1.Evaluate() < Value2.Evaluate())
End Function

Public Function GreaterOrEqual(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set GreaterOrEqual = VBFPValue(Value1.Evaluate() >= Value2.Evaluate())
End Function

Public Function LessOrEqual(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set LessOrEqual = VBFPValue(Value1.Evaluate() <= Value2.Evaluate())
End Function


' ============================================================
' Boolean
' ============================================================

Public Function AndAlso(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set AndAlso = VBFPValue(CBool(Value1.Evaluate()) And CBool(Value2.Evaluate()))
End Function

Public Function OrElse(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set OrElse = VBFPValue(CBool(Value1.Evaluate()) Or CBool(Value2.Evaluate()))
End Function

Public Function NotValue(ByVal Value As IFunction) As IFunction
    Set NotValue = VBFPValue(Not CBool(Value.Evaluate()))
End Function


' ============================================================
' Conditional
' ============================================================

Public Function IfThenElse(ByVal Condition As IFunction, ByVal WhenTrue As IFunction, ByVal WhenFalse As IFunction) As IFunction
    If CBool(Condition.Evaluate()) Then
        Set IfThenElse = WhenTrue
    Else
        Set IfThenElse = WhenFalse
    End If
End Function

' ============================================================
' Strings
' ============================================================

Public Function Concatenate(ByVal Value1 As IFunction, ByVal Value2 As IFunction) As IFunction
    Set Concatenate = VBFPValue(CStr(Value1.Evaluate()) & CStr(Value2.Evaluate()))
End Function

Public Function Length(ByVal Value As IFunction) As IFunction
    Set Length = VBFPValue(Len(Value.Evaluate()))
End Function

Public Function UpperCase(ByVal Value As IFunction) As IFunction
    Set UpperCase = VBFPValue(UCase$(Value.Evaluate()))
End Function

Public Function LowerCase(ByVal Value As IFunction) As IFunction
    Set LowerCase = VBFPValue(LCase$(Value.Evaluate()))
End Function

Public Function TrimValue(ByVal Value As IFunction) As IFunction
    Set TrimValue = VBFPValue(Trim$(Value.Evaluate()))
End Function


' ============================================================
' Collections / Arrays
' ============================================================

Public Function ArrayLength(ByVal Elements As IFunction) As IFunction
    Dim Values() As IFunction
    Values = Elements.Evaluate()

    Set ArrayLength = VBFPValue(UBound(Values) + 1)
End Function

Public Function First(ByVal Elements As IFunction) As IFunction
    Dim Values() As IFunction
    Values = Elements.Evaluate()

    Set First = Values(0)
End Function

Public Function Last(ByVal Elements As IFunction) As IFunction
    Dim Values() As IFunction
    Values = Elements.Evaluate()

    Set Last = Values(UBound(Values))
End Function


' ============================================================
' Predicates
' ============================================================

Public Function IsZero(ByVal Value As IFunction) As IFunction
    Set IsZero = VBFPValue(Value.Evaluate() = 0)
End Function

Public Function IsPositive(ByVal Value As IFunction) As IFunction
    Set IsPositive = VBFPValue(Value.Evaluate() > 0)
End Function

Public Function IsNegative(ByVal Value As IFunction) As IFunction
    Set IsNegative = VBFPValue(Value.Evaluate() < 0)
End Function

Public Function IsEmptyString(ByVal Value As IFunction) As IFunction
    Set IsEmptyString = VBFPValue(Len(CStr(Value.Evaluate())) = 0)
End Function
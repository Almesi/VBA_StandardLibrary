VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "BoundFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IFunction

Private p_Function As IFunction
Private p_Bound()  As Variant

Public Function Create(ByVal n_Function As IFunction, ParamArray Arguments() As Variant) As BoundFunction
    Dim Temp()  As Variant
    Temp = Arguments
    Set Create = CreateArr(n_Function, Temp)
End Function

Public Function CreateArr(ByVal n_Function As IFunction, ByRef Arguments() As Variant) As BoundFunction
    Set CreateArr = New BoundFunction
    Call CreateArr.Initialize(n_Function, Arguments)
End Function

Friend Sub Initialize(ByVal n_Function As IFunction, ByRef n_Bound() As Variant)
    Set p_Function = n_Function
    p_Bound        = n_Bound
End Sub

Private Function IFunction_Evaluate(ParamArray Arguments() As Variant) As Variant
    Dim Temp() As Variant
    Temp = Arguments
    Call VBFPAssign(IFunction_Evaluate, IFunction_EvaluateArr(Temp))
End Function

Private Function IFunction_EvaluateArr(ByRef Arguments() As Variant) As Variant
    Call VBFPAssign(IFunction_EvaluateArr, IFunction_RunArr(Arguments).Evaluate())
End Function

Private Function IFunction_Run(ParamArray Arguments() As Variant) As IFunction
    Dim Temp() As Variant
    Temp = Arguments
    Call VBFPAssign(IFunction_Run, IFunction_RunArr(Temp))
End Function

Private Function IFunction_RunArr(ByRef Arguments() As Variant) As IFunction
    Call VBFPAssign(IFunction_RunArr, Run(Arguments))
End Function

Private Function Run(ByRef Arguments() As Variant) As IFunction
    Dim Combined() As Variant
    Combined = Combine(p_Bound, Arguments)

    Set Run = p_Function.RunArr(Combined)
End Function

Private Function Combine(ByRef First() As Variant, ByRef Second() As Variant) As Variant()
    Dim FirstCount  As Long
    Dim SecondCount As Long
    Dim TotalCount  As Long

    FirstCount  = USize(First)
    SecondCount = USize(Second)

    If FirstCount < 0 And SecondCount < 0 Then
        Dim Emptyy() As Variant
        Combine = Emptyy
        Exit Function
    End If

    TotalCount = (FirstCount + 1) + (SecondCount + 1)

    Dim Result() As Variant
    ReDim Result(TotalCount - 1)

    Dim i As Long
    For i = 0 To FirstCount
        Call VBFPAssign(Result(i), First(i))
    Next i

    Dim Position As Long
    Position = i
    For i = 0 To SecondCount
        Call VBFPAssign(Result(Position + i), Second(i))
    Next i

    Combine = Result
End Function
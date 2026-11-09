VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CurriedFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IFunction

Private p_Function As IFunction
Private p_Argument As IFunction

Friend Sub Initialize(ByVal n_Function As IFunction, ByVal n_Argument As Variant)
    Set p_Function = n_Function
    Call VBFPAssign(p_Argument, n_Argument)
End Sub

Public Function Create(ByVal n_Function As IFunction, ParamArray n_Arguments() As Variant) As CurriedFunction
    Dim Temp() As Variant
    Temp = n_Arguments
    Set Create = CreateArr(n_Function, Temp)
End Function

Public Function CreateArr(ByVal n_Function As IFunction, ByRef n_Arguments() As Variant) As CurriedFunction
    Dim Result As New CurriedFunction
    If USize(n_Arguments) < 0 Then
        Call Result.Initialize(n_Function, Nothing)
        Set CreateArr = Result
        Exit Function
    End If
    Call Result.Initialize(CreateArr(n_Function, Tail(n_Arguments)), n_Arguments(0))
    Set CreateArr = Result
End Function

Private Function IFunction_Evaluate(ParamArray Arguments() As Variant) As Variant
    Dim Temp() As Variant
    Temp = Arguments
    Call VBFPAssign(IFunction_Evaluate, Run(Temp).Evaluate())
End Function

Private Function IFunction_EvaluateArr(ByRef Arguments() As Variant) As Variant
    Call VBFPAssign(IFunction_EvaluateArr, Run(Arguments).Evaluate())
End Function

Private Function IFunction_Run(ParamArray Arguments() As Variant) As IFunction
    Dim Temp() As Variant
    Temp = Arguments
    Call VBFPAssign(IFunction_Run, Run(Temp))
End Function

Private Function IFunction_RunArr(ByRef Arguments() As Variant) As IFunction
    Call VBFPAssign(IFunction_RunArr, Run(Arguments))
End Function

Private Function Run(ByRef Arguments() As Variant) As IFunction
    If p_Argument Is Nothing Then
        Set Run = BoundFunction.CreateArr(p_Function, Arguments)
        Exit Function
    End If

    Dim Combined() As Variant
    Combined = Combine(Arguments)
    Set Run = p_Function.RunArr(Combined)
End Function

Private Function Combine(ByRef Arguments() As Variant) As Variant()
    Dim ArgumentsCount As Long
    ArgumentsCount = USize(Arguments) + 1

    Dim Result() As Variant
    ReDim Result(ArgumentsCount)
    Dim i As Long

    For i = 0 To ArgumentsCount - 1
        Set Result(i) = Arguments(i)
    Next i
    Set Result(ArgumentsCount) = p_Argument
    Combine = Result
End Function

Private Function Tail(ByRef Arr() As Variant) As Variant()
    Dim Size As Long
    Size = USize(Arr) - 1

    Dim Result() As Variant
    If Size < 0 Then
        Tail = Result
        Exit Function
    End If

    ReDim Result(Size)
    Dim i As Long
    For i = 0 To Size
        Set Result(i) = Arr(i + 1)
    Next i
End Function
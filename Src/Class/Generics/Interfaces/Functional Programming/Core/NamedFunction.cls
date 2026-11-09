VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NamedFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IFunction

Private FunctionName      As String
Private ExpectIFunction   As Boolean
Private ArgIsIFunction    As Boolean

Public Function Create(ByVal FunctionName As String, Optional ByVal n_ExpectIFunction As Boolean = False, Optional ByVal n_ArgIsIFunction As Boolean = False) As NamedFunction
    Set Create = New NamedFunction
    Call Create.Initialize(FunctionName, n_ExpectIFunction, n_ArgIsIFunction)
End Function

Friend Sub Initialize(ByVal n_Name As String, Optional ByVal n_ExpectIFunction As Boolean = False, Optional ByVal n_ArgIsIFunction As Boolean = False)
    Let FunctionName    = n_Name
    Let ExpectIFunction = n_ExpectIFunction
    Let ArgIsIFunction = n_ArgIsIFunction
End Sub

Private Function IFunction_Evaluate(ParamArray Arguments() As Variant) As Variant
    Dim Temp() As Variant

    Temp = Arguments
    Call VBFPAssign(IFunction_Evaluate, IFunction_RunArr(Temp).Evaluate())
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
    Dim Passed() As Variant

    Passed = FuncArgumentsToArguments(Arguments, ExpectIFunction, ArgIsIFunction)
    If ArgIsIFunction Then
        Call VBFPAssign(IFunction_RunArr, RunFunction(Passed))
    Else
        Call VBFPAssign(IFunction_RunArr, ValueFunction.Create(RunFunction(Passed)))
    End If
End Function

Private Function RunFunction(ByRef A() As Variant) As Variant
    Dim Count As Long
    Count = USize(A)
    Select Case Count + 1
        Case 00: Call VBFPAssign(RunFunction, Application.Run(FunctionName))
        Case 01: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0)))
        Case 02: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1)))
        Case 03: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2)))
        Case 04: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3)))
        Case 05: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4)))
        Case 06: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5)))
        Case 07: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6)))
        Case 08: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7)))
        Case 09: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8)))
        Case 10: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9)))
        Case 11: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10)))
        Case 12: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11)))
        Case 13: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12)))
        Case 14: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13)))
        Case 15: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14)))
        Case 16: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15)))
        Case 17: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16)))
        Case 18: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17)))
        Case 19: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18)))
        Case 20: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19)))
        Case 21: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20)))
        Case 22: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21)))
        Case 23: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22)))
        Case 24: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23)))
        Case 25: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24)))
        Case 26: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25)))
        Case 27: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25), A(26)))
        Case 28: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25), A(26), A(27)))
        Case 29: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25), A(26), A(27), A(28)))
        Case 30: Call VBFPAssign(RunFunction, Application.Run(FunctionName, A(0), A(1), A(2), A(3), A(4), A(5), A(6), A(7), A(8), A(9), A(10), A(11), A(12), A(13), A(14), A(15), A(16), A(17), A(18), A(19), A(20), A(21), A(22), A(23), A(24), A(25), A(26), A(27), A(28), A(29)))
        Case Else: Err.Raise 5, "NamedFunction", "Unsupported number of arguments: " & Count
    End Select
End Function
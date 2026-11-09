VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ComposedFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IFunction

Private p_Outer As IFunction
Private p_Inner As IFunction

Friend Sub Initialize(ByVal n_Outer As IFunction, ByVal n_Inner As IFunction)
    Set p_Outer = n_Outer
    Set p_Inner = n_Inner
End Sub

Public Function Create(ByVal n_Outer As IFunction, ByVal n_Inner As IFunction) As ComposedFunction
    Set Create = New ComposedFunction
    Call Create.Initialize(n_Outer, n_Inner)
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
    Set Run = p_Outer.Run(p_Inner.RunArr(Arguments))
End Function
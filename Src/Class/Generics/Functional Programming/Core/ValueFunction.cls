VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IFunction

Private p_Value As Variant

Friend Sub Initialize(ByVal n_Value As Variant)
    If IsObject(n_Value) Then
        Set p_Value = n_Value
    Else
        Let p_Value = n_Value
    End If
End Sub

Public Function Create(ByVal Value As Variant) As ValueFunction
    Set Create = New ValueFunction
    Call Create.Initialize(Value)
End Function

Private Function IFunction_Evaluate(ParamArray Arguments() As Variant) As Variant
    Call VBFPAssign(IFunction_Evaluate, p_Value)
End Function

Private Function IFunction_EvaluateArr(ByRef Arguments() As Variant) As Variant
    Call VBFPAssign(IFunction_Evaluate, p_Value)
End Function

Private Function IFunction_Run(ParamArray Arguments() As Variant) As IFunction
    Set IFunction_Run = Me
End Function

Private Function IFunction_RunArr(ByRef Arguments() As Variant) As IFunction
    Set IFunction_Run = Me
End Function
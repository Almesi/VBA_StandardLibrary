Attribute VB_Name = "EventHandler"

Option Explicit

Public Function RunEvent(ByVal FunctionsDictionary As FunctionDir, ByVal Key As String, ParamArray Arguments() As Variant) As Variant
    If FunctionsDictionary Is Nothing Then Exit Function
    
    Dim Found As IFunction
    Set Found = FunctionsDictionary.Func(Key)

    If Found Is Nothing Then Exit Function

    Dim Args() As Variant
    Args = Arguments
    Call VBFPAssign(RunEvent, Found.EvaluateArr(Args))
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "MemoizedFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IFunction

Private p_Function As IFunction
Private p_Cache    As Object


' Root of the memoization tree.
'
' Each node is a Dictionary:
'
'   Argument 1
'       |
'       +-- Argument 2
'               |
'               +-- Argument 3
'                       |
'                       +-- Result
'

Private Sub Class_Initialize()
    Set p_Cache = CreateObject("Scripting.Dictionary")
End Sub

Friend Sub Initialize(ByVal n_Function As IFunction)
    Set p_Function = n_Function
End Sub

Public Function Create(ByVal n_Function As IFunction) As MemoizedFunction
    Set Create = New MemoizedFunction
    Call Create.Initialize(n_Function)
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
    Dim Current As Object
    Dim Key     As String
    Dim i       As Long
    Dim Count   As Long

    Set Current = p_Cache
    Count = USize(Arguments)
    For i = 0 To Count
        Key = ArgumentKey(Arguments(i))

        If Not Current.Exists(Key) Then
            Call Current.Add(Key, CreateObject("Scripting.Dictionary"))
        End If

        Set Current = Current(Key)
    Next i

    If Current.Exists(ResultKey()) Then
        Set Run = Current(ResultKey())
    Else
        Set Run = p_Function.RunArr(Arguments)
        Call Current.Add(ResultKey(), Run)
    End If
End Function

Private Function ResultKey() As String
    ResultKey = ChrW(0) & "RESULT" & ChrW(0)
End Function

Private Function ArgumentKey(ByVal Argument As Variant) As String
    On Error GoTo Error
    Dim Temp As IFunction
    Set Temp = Argument

    ArgumentKey = ValueKey(Temp.Evaluate())
    Exit Function

    Error:
    ArgumentKey = ValueKey(Argument)
End Function

Private Function ValueKey(ByVal Value As Variant) As String
    Dim Result As String
    Dim i As Long

    Result = CStr(VarType(Value)) & ":"
    If IsObject(Value) Then
        Result = Result & "OBJECT:" & ObjectKey(Value)
    ElseIf IsArray(Value) Then
        Result = Result & "ARRAY:"

        On Error GoTo NotAnArray
        For i = 0 To UBound(Value)
            Result = Result & _
                     Len(ValueKey(Value(i))) & ":" & _
                     ValueKey(Value(i))
        Next i
        Exit Function

    NotAnArray:
        Result = Result & "EMPTY"
    Else
        Select Case VarType(Value)
            Case vbEmpty   : Result = Result & "EMPTY"
            Case vbNull    : Result = Result & "NULL"
            Case vbString  : Result = Result & Len(CStr(Value)) & ":" & CStr(Value)
            Case vbBoolean : Result = Result & CStr(CBool(Value))
            Case vbDate    : Result = Result & CStr(CDbl(CDate(Value)))
            Case Else      : Result = Result & CStr(Value)
        End Select
    End If
    ValueKey = Result
End Function

Private Function ObjectKey(ByVal Value As Object) As String
    ObjectKey = TypeName(Value) & ":" & CStr(ObjPtr(Value))
End Function
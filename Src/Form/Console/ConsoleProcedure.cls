VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
 

Private p_Name As String
Private p_ReturnType As String
Private p_Arguments As String
Private p_Parent As String
Private p_Value As Variant

Public Property Let Name(n_Name As String)             : p_Name       = n_Name         : End Property
Public Property Let ReturnType(n_ReturnType As String) : p_ReturnType = n_ReturnType   : End Property
Public Property Let Arguments(n_Arguments As String)   : p_Arguments  = n_Arguments    : End Property
Public Property Let Parent(n_Parent As String)         : p_Parent     = n_Parent       : End Property
Public Property Let Value(n_Value As Variant)          : Call Assign(p_Value, n_Value) : End Property

Public Property Get Name() As String                   : Name         = p_Name         : End Property
Public Property Get ReturnType() As String             : ReturnType   = p_ReturnType   : End Property
Public Property Get Parent() As String                 : Parent       = p_Parent       : End Property
Public Property Get Arguments() As String              : Arguments    = p_Arguments    : End Property
Public Property Get Value() As Variant                 : Call Assign(Value, p_Value)   : End Property

Public Function CreateProcedure(n_Name As String, n_ReturnType As String, n_Arguments As String, n_Parent As String) As ConsoleProcedure
    Set CreateProcedure = New ConsoleProcedure
    With CreateProcedure
        .Name       = n_Name
        .ReturnType = n_ReturnType
        .Arguments  = n_Arguments
        .Parent     = n_Parent
        .Value      = Empty
    End With
End Function

Public Function CreateVariable(n_Name As String, n_Parent As String, n_Value As Variant) As ConsoleProcedure
    Set CreateVariable = New ConsoleProcedure
    With CreateVariable
        .Name       = n_Name
        .ReturnType = VarType(n_Value)
        .Arguments  = Empty
        .Parent     = n_Parent
        .Value      = n_Value
    End With
End Function

Public Function Create(n_Name As String, n_ReturnType As String, n_Arguments As String, n_Parent As String, n_Value As Variant) As ConsoleProcedure
    Set Create = New ConsoleProcedure
    With Create
        .Name       = n_Name
        .ReturnType = n_ReturnType
        .Arguments  = n_Arguments
        .Parent     = n_Parent
        .Value      = n_Value
    End With
End Function

Public Function CreateLambda(n_Name As String, n_Parent As String, Optional Equation As Variant = "", Optional UsePerformanceCache As Boolean = False, Optional SandboxExtras As Boolean = False) As ConsoleProcedure
    Dim Temp As stdLambda
    Set CreateLambda = New ConsoleProcedure
    Set Temp = stdLambda.Create(Equation, UsePerformanceCache, SandboxExtras)
    With CreateLambda
        .Name       = n_Name
        .ReturnType = TypeName(Temp)
        .Arguments  = "Arguments might exist"
        .Parent     = n_Parent
        .Value      = Temp
    End With
End Function

Public Function CreateScript(n_Name As String, n_Arguments As String, n_Script As String, n_Parent As String) As ConsoleProcedure
    If Arguments = "" Then Arguments = "No Arguments"
    Set CreateScript = New ConsoleProcedure
    With CreateScript
        .Name       = n_Name
        .ReturnType = "ConsoleScript"
        .Arguments  = n_Arguments
        .Parent     = n_Parent
        .Value      = n_Script
    End With
End Function

Public Function Run(ReturnVariable As Variant, Arguments() As Variant) As Boolean
    On Error GoTo Error
    Select Case ReturnType
        Case "stdICallable" : Assign(ReturnVariable, Value.RunEx(Arguments))
        Case Else           : Assign(ReturnVariable, Value)
    End Select
    Run = True
    Exit Function

    Error:
End Function

Private Sub Assign(Goal As Variant, Value As Variant)
    If IsObject(Value) Then
        Set Goal = Value
    Else
        Goal = Value
    End If
End Sub
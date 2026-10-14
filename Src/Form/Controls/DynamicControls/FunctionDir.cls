VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "FunctionDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Just to explicitly use a class that is not ambigous

Private pFunctions As Object ' Dictionary

Private Sub Class_Initialize()
    Set pFunctions = CreateObject("Scripting.Dictionary")
End Sub

Public Function Create() As FunctionDir
    Set Create = New FunctionDir
End Function

Public Function Exists(ByVal Name As String) As Boolean
    Exists = pFunctions.Exists(Name)
End Function

Public Property Get Func(ByVal Name As String) As IFunction
    If Exists(Name) Then Set Func = pFunctions(Name)
End Property

Public Sub Add(ByVal Name As String, ByVal NewFunc As IFunction)
    If pFunctions.Exists(Name) Then
        Set pFunctions(Name) = NewFunc
    Else
        Call pFunctions.Add(Name, NewFunc)
    End If
End Sub
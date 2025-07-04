VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextColoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private p_ColorDef As ColorDefinition
Private p_Rules() As stdLambda

Public Property Let ColorDef(n_ColorDef As ColorDefinition)
    Dim Temp() As String
    Set p_ColorDef = n_ColorDef
    Temp = n_ColorDef.Names
    ReDim p_Rules(Ubound(Temp))
End Property
Public Property Get ColorDef() As ColorDefinition
    Set ColorDef = p_ColorDef
End Property


Public Property Let Rule(Index As Long, Value As stdLambda)
    Set p_Rules(Index) = Value
End Property
Public Property Get Rule(Index As Long) As stdLambda
    Set Rule = p_Rules(Index)
End Property

Public Sub MakeRule(Name As String, RuleText As String)
    Dim i As Long
    For i = 0 To Ubound(ColorDef.Names)
        If ColorDef.Name(i) = Name Then
            Call MakeRuleByIndex(i, RuleText)
        End If
    Next i
End Sub
Public Sub MakeRuleByIndex(Index As Long, RuleText As String)
    Dim Temp As stdLambda
    Set Temp = stdLambda.Create(RuleText)
    Rule(Index) = Temp
End Sub

Public Function GetColor(Text As String) As Long
    Dim i As Long
    For i = 0 To Ubound(p_Rules)
        If p_Rules(i).Run(Text) Then
            GetColor = ColorDef.ColorByIndex(i)
            Exit Function
        End If
    Next i
End Function







Public Property Let Basic(n_Basic             As Long) : p_Basic       = n_Basic       : End Property
Public Property Let System(n_System           As Long) : p_System      = n_System      : End Property
Public Property Let Procedure(n_Procedure     As Long) : p_Procedure   = n_Procedure   : End Property
Public Property Let Operator(n_Operator       As Long) : p_Operator    = n_Operator    : End Property
Public Property Let Datatype(n_Datatype       As Long) : p_Datatype    = n_Datatype    : End Property
Public Property Let Value(n_Value             As Long) : p_Value       = n_Value       : End Property
Public Property Let Stringg(n_String          As Long) : p_String      = n_String      : End Property
Public Property Let Statement(n_Statement     As Long) : p_Statement   = n_Statement   : End Property
Public Property Let Keyword(n_Keyword         As Long) : p_Keyword     = n_Keyword     : End Property
Public Property Let Parantheses(n_Parantheses As Long) : p_Parantheses = n_Parantheses : End Property
Public Property Let Variable(n_Variable       As Long) : p_Variable    = n_Variable    : End Property
Public Property Let Script(n_Script           As Long) : p_Script      = n_Script      : End Property
Public Property Let Lambda(n_Lambda           As Long) : p_Lambda      = n_Lambda      : End Property

Public Property Get Basic()                   As Long  : Basic         = p_Basic       : End Property
Public Property Get System()                  As Long  : System        = p_System      : End Property
Public Property Get Procedure()               As Long  : Procedure     = p_Procedure   : End Property
Public Property Get Operator()                As Long  : Operator      = p_Operator    : End Property
Public Property Get Datatype()                As Long  : Datatype      = p_Datatype    : End Property
Public Property Get Value()                   As Long  : Value         = p_Value       : End Property
Public Property Get Stringg()                 As Long  : Stringg       = p_String      : End Property
Public Property Get Statement()               As Long  : Statement     = p_Statement   : End Property
Public Property Get Keyword()                 As Long  : Keyword       = p_Keyword     : End Property
Public Property Get Parantheses()             As Long  : Parantheses   = p_Parantheses : End Property
Public Property Get Variable()                As Long  : Variable      = p_Variable    : End Property
Public Property Get Script()                  As Long  : Script        = p_Script      : End Property
Public Property Get Lambda()                  As Long  : Lambda        = p_Lambda      : End Property
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Interpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

' needs to be able to take a text as input and correctly interpret it
' as stdICallable
' needs to know if it is an existing macro or procedure

Private Const ArgSeperator          As String = ", " 
Private Const AsgOperator           As String = " = "
Private Const LineSeperator         As String = "LINEBREAK/()\"

Private       Errors                As ArrayValueModel
Private       PublicProcedures      As ArrayValueModel
Private       PrivateProcedures     As ArrayValueModel
Private       Publics()             As ArrayValueModel


Private NoArguments() As Variant



Public Property Get ErrorMessages() As ArrayValueModel
    Set ErrorMessages = Errors
End Property

Public Property Get PublicVar() As ArrayValueModel
    Set PublicVar = PublicProcedures
End Property
Public Property Let PublicVar(n_Value As ArrayValueModel)
    Set PublicProcedures = n_Value
End Property

Public Property Get PrivateVar() As ArrayValueModel
    Set PrivateVar = PrivateProcedures
End Property
Public Property Let PrivateVar(n_Value As ArrayValueModel)
    Set PrivateProcedures = n_Value
End Property

Private Sub Class_Initialize()
    Set Errors           = New ArrayValueModel
    Set PublicVariables  = New ArrayValueModel
    Set PrivateVariables = New ArrayValueModel
    Call GetAllWorkbooksProcedures()
End Sub

Private Sub Class_Terminate()
End Sub

Public Function Run(ReturnVariable As Variant, Line As String, Optional DimVariables As ArrayValueModel) As Boolean
    Dim AssignOperator  As Long
    Dim RightSide       As String
    Dim LeftSide        As String
    Dim Procedure       As ConsoleProcedure

    AssignOperator = InStr(1, Line, AsgOperator)
    If AssignOperator <> 0 Then LeftSide = MidP(Line, 1, AssignOperator - 1)
    If InStr(1, Line, AsgOperator) <> 0 Then
        RightSide = MidP(Line, AssignOperator + Len(AsgOperator), Len(Line))
    Else
        RightSide = MidP(Line, 1, Len(Line))
    End If
    
    If RunProcedure(ReturnVariable, RightSide) = False Then
        Select Case True
            Case IsNumeric(RightSide), IsDate(RightSide)
                Run = InterpretVariable(ReturnVariable, RightSide)
            Case InString(RightSide, 1, Len(RightSide))
                ReturnVariable = GetStringString(RightSide)
            Case Mid(RightSide, 1, 1) Like "[?]*"
                Set ReturnVariable = ConsoleProcedure.CreateLambda(LeftSide, Mid(RightSide, 2, Len(RightSide)), False, False)
            Case Else
                Run = HandleSpecial(ReturnVariable, RightSide)
                If Run = False Then
                    Run = HandleReturnOperator(ReturnVariable, RightSide)
                    If Run = False Then
                        Run = RunProcedure(ReturnVariable, RightSide, NoArguments)
                    End If
                End If
        End Select
    End If

    If LeftSide <> Empty Then
        Dim LeftSideModel As ArrayValueModel
        Dim LeftSideModelIndex As Long
        Dim Procedure As ConsoleProcedure
        LeftSideModelIndex = VariableIndex(LeftSideModel, LeftSide, DimVariables)
        If LeftSideModel Is Nothing Then
            If TypeName(ReturnVariable) = TypeName(New ConsoleProcedure) Then
                LeftSideModelIndex = AddNewVariable(LeftSideModel, "PRIVATE " & LeftSide & " As " & ReturnVariable.ReturnType)
                Set Procedure = LeftSideModel.Element(LeftSideModelIndex)
                Procedure.Value = ReturnVariable.Value
            Else
                LeftSideModelIndex = AddNewVariable(LeftSideModel, "PRIVATE " & LeftSide & " As " & TypeName(ReturnVariable))
                Set Procedure = LeftSideModel.Element(LeftSideModelIndex)
                Procedure.Value = ReturnVariable
            End If
        End If
        Set Procedure = LeftSideModel.Element(LeftSideModelIndex)
        If TypeName(ReturnVariable) = TypeName(New ConsoleProcedure) Then
            Procedure.Value = ReturnVariable.Value
        Else
            Procedure.Value = ReturnVariable
        End If
        LeftSideModel.Element(LeftSideModelIndex) = Procedure
        ReturnVariable = Empty
    End If
End Function

Private Function RunProcedure(ReturnVariable As Variant, Text As String) As Boolean
    Dim Procedure As ConsoleProcedure
    Dim Name As String
    Dim Arguments() As Variant

    Name = GetProcedureName(Text)
    Arguments = HandleArguments(Text)
    Set Procedure = GetProcedureInPublics(Name)
    If Procedure Is Nothing Then Set Procedure = GetProcedure(PublicProcedures, Name)
    If Procedure Is Nothing Then Set Procedure = GetProcedure(PrivateProcedures, Name)
    If Not Procedure Is Nothing Then
        RunProcedure = Procedure.Run(ReturnVariable, Arguments)
        Exit Function
    End If
End Function

Public Sub GetAllWorkbooksProcedures()
    Dim WB As Workbook
    Dim i As Long
    ReDim Publics(WorkBooks.Count - 1)
    For Each WB in WorkBooks
        Call ExpandPublicProcedures(i, WB)
        i = i + 1
    Next WB
End Sub

Public Sub ExpandPublicProcedures(Index As Long, WB As Workbook)
    Dim VBComp As VBIDE.VBComponent
    For Each VBComp In WB.VBProject.VBComponents
        Call GetPublicsInComponent(Publics(Index), VBComp)
    Next VBComp
End Sub

Private Function GetProcedureInPublics(Name As String) As ConsoleProcedure
    Dim i As Long
    For i = 0 To ArraySize(Publics)
        Set GetProcedureInPublics = GetProcedure(Publics(i), Name)
        If Not GetProcedureInPublics Is Nothing Then Exit Sub
    Next i
End Function

Private Function GetProcedure(ByRef Model As ArrayValueModel, Name As String) As ConsoleProcedure
    Dim i As Long
    For i = 0 To Model.Count
        If Model.Element.Name = Name Then
            Set GetProcedure = Model.Element
        End If
    Next i
End Function

Private Sub GetPublicsInComponent(Model As ArrayValueModel, VBComp As VBIDE.VBComponent)
    Dim CodeMod As VBIDE.CodeModule
    Dim CurrentProcedure As ConsoleProcedure
    Dim Line As String
    Dim TempArray() As String
    Dim LastWord As String
    Dim Value As stdICallable
    Dim mType As String

    mType = IIf(VBComp.Type = vbext_ComponentType.vbext_ct_StdModule, "Module", "Object")
    For i = 1 To CodeMod.CountOfLines
        Line = CodeMod.Lines(i, 1)
        If UCase(Line) Like "*PUBLIC *" And InStr(1, Line, "'") = 0 And Not UCase(Line) Like "*" & Chr(34) & "*PUBLIC*" & Chr(34) & "*" Then
            Set CurrentProcedure = New ConsoleProcedure
            TempArray = Split(Line, " ")
            LastWord = TempArray(ArraySize(TempArray))
            If IsProcedure(Line) Then
                CurrentProcedure.Name      = ReturnProcedure(Line)
                CurrentProcedure.Arguments = GetParanthesesText(Line)
                If Mid(LastWord, Len(LastWord), 1) = ")" Then
                    CurrentProcedure.ReturnType = "Void"
                Else
                    CurrentProcedure.ReturnType = LastWord
                End If
                ' Probably needs to accomadate for class object and not just public procedures
                Set Value = stdCallback.Create(mType, VBComp.Name, CurrentProcedure.Name, GetCallType(Line))
                CurrentProcedure.Value = Value
            Else
                CurrentProcedure.Name = ReturnVariable(Line)
                CurrentProcedure.Arguments = Empty
                CurrentProcedure.ReturnType = LastWord
                CurrentProcedure.Value = Empty
            End If
            Call Model.Add(CurrentProcedure)
        End If
    Next i
End Sub

Private Function IsProcedure(Line As String) As Boolean
    If (UCase(Line) Like "* FUNCTION *" Or _
        UCase(Line) Like "* SUB *"      Or _
        UCase(Line) Like "* SET *"      Or _
        UCase(Line) Like "* LET *"      Or _
        UCase(Line) Like "* GET *"      Or _
        UCase(Line) Like "* EVENT *") Then
End Function

'                          |----------|
'   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
Private Function ReturnProcedure(Line As String) As String
    Dim StartPoint As Long
    Dim EndPoint As Long
    Select Case True
        Case UCase(Line) Like "*PUBLIC STATIC SUB *(*)*"      : StartPoint = InStr(1, UCase(Line), "PUBLIC STATIC SUB ")      + Len("PUBLIC STATIC SUB ")      : EndPoint = InStr(1, UCase(Line), "(")
        Case UCase(Line) Like "*PUBLIC SUB *(*)*"             : StartPoint = InStr(1, UCase(Line), "PUBLIC SUB ")             + Len("PUBLIC SUB ")             : EndPoint = InStr(1, UCase(Line), "(")
        Case UCase(Line) Like "*PUBLIC STATIC FUNCTION *(*)*" : StartPoint = InStr(1, UCase(Line), "PUBLIC STATIC FUNCTION ") + Len("PUBLIC STATIC FUNCTION ") : EndPoint = InStr(1, UCase(Line), "(")
        Case UCase(Line) Like "*PUBLIC FUNCTION *(*)*"        : StartPoint = InStr(1, UCase(Line), "PUBLIC FUNCTION ")        + Len("PUBLIC FUNCTION ")        : EndPoint = InStr(1, UCase(Line), "(")
        Case UCase(Line) Like "*PUBLIC EVENT *(*)*"           : StartPoint = InStr(1, UCase(Line), "PUBLIC EVENT ")           + Len("PUBLIC EVENT ")           : EndPoint = InStr(1, UCase(Line), "(")
        Case UCase(Line) Like "*PUBLIC PROPERTY GET *(*)*"    : StartPoint = InStr(1, UCase(Line), "PUBLIC PROPERTY GET ")    + Len("PUBLIC PROPERTY GET ")    : EndPoint = InStr(1, UCase(Line), "(")
        Case UCase(Line) Like "*PUBLIC PROPERTY SET *(*)*"    : StartPoint = InStr(1, UCase(Line), "PUBLIC PROPERTY SET ")    + Len("PUBLIC PROPERTY SET ")    : EndPoint = InStr(1, UCase(Line), "(")
        Case UCase(Line) Like "*PUBLIC PROPERTY LET *(*)*"    : StartPoint = InStr(1, UCase(Line), "PUBLIC PROPERTY LET ")    + Len("PUBLIC PROPERTY LET ")    : EndPoint = InStr(1, UCase(Line), "(")
        Case Else
            Exit Function
    End Select
    ReturnProcedure = MidP(Line, StartPoint, EndPoint - 1)
End Function

'          |----------|
'   Public VariableName As Variant
Private Function ReturnVariable(Line As String) As String
    Dim StartPoint As Long
    Dim EndPoint As Long
    Select Case True
        Case UCase(Line) Like "*PUBLIC CONST *": StartPoint = InStr(1, UCase(Line), "*PUBLIC CONST *") + Len("PUBLIC CONST "): EndPoint = InStr(1, UCase(Line), " AS ")
        Case UCase(Line) Like "*PUBLIC ENUM *":  StartPoint = InStr(1, UCase(Line), "*PUBLIC ENUM *")  + Len("PUBLIC ENUM "):  EndPoint = Len(Line)                    
        Case UCase(Line) Like "*PUBLIC TYPE *":  StartPoint = InStr(1, UCase(Line), "*PUBLIC TYPE *")  + Len("PUBLIC TYPE "):  EndPoint = Len(Line)                    
        Case UCase(Line) Like "*PUBLIC *":       StartPoint = InStr(1, UCase(Line), "*PUBLIC *")       + Len("PUBLIC "):       EndPoint = InStr(1, UCase(Line), " AS ")
        Case Else
            Exit Function
    End Select
    ReturnVariable =  MidP(Line, StartPoint + 1, EndPoint)
End Function

Private Function GetCallType(Text As String) As VBCallType
    Select Case True
        Case UCase(Line) Like "*PUBLIC STATIC SUB *(*)*"      : GetCallType = VBCallType.VBMethod
        Case UCase(Line) Like "*PUBLIC SUB *(*)*"             : GetCallType = VBCallType.VBMethod
        Case UCase(Line) Like "*PUBLIC STATIC FUNCTION *(*)*" : GetCallType = VBCallType.VBMethod
        Case UCase(Line) Like "*PUBLIC FUNCTION *(*)*"        : GetCallType = VBCallType.VBMethod
        Case UCase(Line) Like "*PUBLIC EVENT *(*)*"           : GetCallType = VBCallType.VBMethod
        Case UCase(Line) Like "*PUBLIC PROPERTY GET *(*)*"    : GetCallType = VBCallType.VBGet
        Case UCase(Line) Like "*PUBLIC PROPERTY SET *(*)*"    : GetCallType = VBCallType.VBSet
        Case UCase(Line) Like "*PUBLIC PROPERTY LET *(*)*"    : GetCallType = VBCallType.VBLet
        Case Else
    End Select
End Function

Private Sub AddToArray(Arr As Variant, Value As Variant)
    ReDim Preserve Arr(ArraySize(Arr) + 1)
    Call Assign(Arr(ArraySize), Value)
End Sub

Private Function ArraySize(Arr As Variant, Optional Dimension As Long = 1) As Long
    On Error Resume Next
    ArraySize = -1
    ArraySize = Ubound(Arr, Dimension)
End Function

Private Sub Assign(Goal As Variant, Value As Variant)
    If IsObject(Value) Then
        Set Goal = Value
    Else
        Goal = Value
    End If
End Sub
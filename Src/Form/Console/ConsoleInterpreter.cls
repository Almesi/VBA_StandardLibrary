VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleInterpreter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit


    Private Const ArgSeperator          As String = ", " 
    Private Const AsgOperator           As String = " = "
    Private Const LineSeperator         As String = "LINEBREAK/()\"

    Private       Errors                As ArrayValueModel
    Private       PublicVariables       As ArrayValueModel
    Private       PrivateVariables      As ArrayValueModel

    Private Publics() As std_Tree
    Private NoArguments() As Variant
'

' Initialization

    Public Property Get PublicVar() As ArrayValueModel
        Set PublicVar = PublicVariables
    End Property
    Public Property Let PublicVar(n_Value As ArrayValueModel)
        Set PublicVariables = n_Value
    End Property

    Public Property Get ErrorMessages() As ArrayValueModel
        Set ErrorMessages = Errors
    End Property

    Public Property Get PrivateVar() As ArrayValueModel
        Set PrivateVar = PrivateVariables
    End Property
    Public Property Let PrivateVar(n_Value As ArrayValueModel)
        Set PrivateVariables = n_Value
    End Property

    Private Sub Class_Initialize()
        Set Errors           = New ArrayValueModel
        Set PublicVariables  = New ArrayValueModel
        Set PrivateVariables = New ArrayValueModel
        Call InitializeTree()
        Call GetAllProcedures()
    End Sub

    Private Sub InitializeTree()
        Dim i As Long
        Dim WB As Workbook
        Dim Temp As ConsoleProcedure
        ReDim Publics(Workbooks.Count - 1)
        i = 0
        For Each WB In Workbooks
            Set Publics(i) = std_Tree.Create()
            Set Temp = ConsoleProcedure.Create(WB.VBProject.Name, "As VBProject", "No Arguments", "No Value")
            Publics(i).Value(0) = Temp
            i = i + 1
        Next
    End Sub

    Private Sub Class_Terminate()
    End Sub
'

' Get/Set Values

    Public Function GetProcedureArguments(Line As String) As Variant()
        Dim Temp() As String
        Temp = Split(GetParanthesesText(Line), ArgSeperator)
        GetProcedureArguments = ArrayConvertVariant(Temp)
    End Function

    Public Function GetProcedurePosition(ReturnTree As std_Tree, ProcedureName As String) As Long
        Dim i As Long, j As Long

        For i = 0 To ArraySize(Publics)
            For j = 0 To ArraySize(Publics(i).Tree)
                If Publics(i).Value(j).Name = ProcedureName Then
                    Set ReturnTree = Publics(i)
                    GetProcedurePosition = j
                    Exit Function
                End If
            Next j
        Next i
        GetProcedurePosition = -1
    End Function

    Public Function VariableIndex(ReturnModel As ArrayValueModel, ProcedureName As String, Optional DimVariables As ArrayValueModel = Nothing) As Long
        Dim i As Long
        If Not DimVariables Is Nothing Then
            For i = 0 To ArraySize(DimVariables.Arr)
                If DimVariables.Element(i).Name = ProcedureName Then
                    Set ReturnModel = DimVariables
                    VariableIndex = i
                    Exit Function
                End If
            Next i
        End If

        For i = 0 To ArraySize(PrivateVariables.Arr)
            If PrivateVariables.Element(i).Name = ProcedureName Then
                Set ReturnModel = PrivateVariables
                VariableIndex = i
                Exit Function
            End If
        Next i

        For i = 0 To ArraySize(PublicVariables.Arr)
            If PublicVariables.Element(i).Name = ProcedureName Then
                Set ReturnModel = PublicVariables
                VariableIndex = i
                Exit Function
            End If
        Next i
        VariableIndex = -1
    End Function

    Public Function GetProcedure(ReturnVariable As ConsoleProcedure, Name As String, Optional DimVariables As ArrayValueModel = Nothing) As Boolean
        Dim ReturnTree As std_Tree
        Dim ReturnIndex As Long
        Dim ReturnModel As ArrayValueModel

        ReturnIndex = GetProcedurePosition(ReturnTree, Name)
        If Not ReturnTree Is Nothing Then
            Set ReturnVariable = ReturnTree.Value(ReturnIndex)
            GetProcedure = True
        Else
            ReturnIndex = VariableIndex(ReturnModel, Name, DimVariables)
            If Not ReturnModel Is Nothing Then
                Set ReturnVariable = ReturnModel.Element(ReturnIndex)
                GetProcedure = True
            End If
        End If
    End Function

    Public Function GetAllOperators(Variable() As Variant) As Variant()

        Dim Operators() As Variant
        Dim FoundOperators() As Variant
        Dim Temp() As Variant
        Dim TempStr() As String
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Operators = Array(" IS ", "==", "<>", "=>", "=<", "<=", ">=", "<", ">", " NOT ", " AND ", " OR ", " XOR ", "!=", "||", "&&", "//", "**", "++", "--", "^^",  "+", "-", "*", "/", "^", "?", ":", ";", "!", "|", "?", ",", "&")
        For i = 0 To ArraySize(Operators)
            For j = 0 To ArraySize(Variable)
                If InStr(1, Variable(j), Operators(i)) > 1 Then
                    For k = 0 To ArraySize(FoundOperators)
                        If InStr(1, FoundOperators(k), Operators(i)) Then GoTo Skip
                    Next
                    FoundOperators = ArrayPush(FoundOperators, Operators(i))
                    TempStr = Split(CStr(Variable(j)), CStr(Operators(i)))
                    ReDim Temp(ArraySize(TempStr))
                    For k = 0 To ArraySize(TempStr)
                        Temp(k) = Replace(TempStr(k), " ", "")
                    Next
                    Temp = ArrayInsertEach(Temp, Operators(i))
                    If (ArraySize(Variable) - 1) <> -1 Then
                        Variable = ArrayRemove(Variable, j)
                        Variable = ArrayInsertArray(Variable, Temp, j)
                    Else
                        Variable = Temp
                    End If
                End If
            Next j
            Skip:
        Next i
        GetAllOperators = Variable

    End Function

    Public Function GetStringString(Text As String) As String
        GetStringString = MidP(Text, 2, Len(Text) - 1)
    End Function

    Public Function HasOperators(Text As String) As Boolean
        Dim i As Long
        Dim Operators As Variant
        Operators = Array(" IS ", "==", "<>", "=>", "=<", "<=", ">=", "<", ">", " NOT ", " AND ", " OR ", " XOR ", "!=", "||", "&&", "//", "**", "++", "--", "^^",  "+", "-", "*", "/", "^", "?", ":", ";", "!", "|", "?", ",", "&")
        For i = 0 To ArraySize(Operators)
            If Instr(1, Text, Operators(i)) <> 0 Then
                HasOperators = True
                Exit Function
            End If
        Next i
    End Function

    Public Function RepeatRun(ReturnVariable As Variant, Lines() As String, Optional DimVariables As ArrayValueModel = Nothing) As Boolean
        Dim i As Long
        Dim NoError As Boolean
        For i = 0 To ArraySize(Lines)
            If NoError Then
                NoError = Run(ReturnVariable, Lines(i), DimVariables)
            Else
                Call      Run(ReturnVariable, Lines(i), DimVariables)
            End If
        Next i
        RepeatRun = NoError
    End Function


'

' Handle Input

    Public Function Run(ReturnVariable As Variant, Line As String, Optional DimVariables As ArrayValueModel) As Boolean
        Dim AssignOperator  As Long
        Dim RightSide       As String
        Dim LeftSide        As String

        AssignOperator = InStr(1, Line, AsgOperator)
        If AssignOperator <> 0 Then LeftSide = Mid(Line, 1, AssignOperator - 1)
        If InStr(1, Line, AsgOperator) <> 0 Then
            RightSide = Mid(Line, AssignOperator + Len(AsgOperator), Len(Line))
        Else
            RightSide = Mid(Line, 1, Len(Line))
        End If

        If RunProcedure(ReturnVariable, RightSide, HandleArguments(RightSide)) = False Then
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

    Private Function RunProcedure(ReturnVariable As Variant, Line As String, Arguments() As Variant) As Boolean
        Dim Name As String
        Name = GetProcedureName(Line)

        Dim Tree As std_Tree
        Dim TreeIndex As Long
        TreeIndex = GetProcedurePosition(Tree, Name)
        If Not Tree Is Nothing Then
            If ArraySize(Arguments) <> -1 Then
                RunProcedure = RunApplication(ReturnVariable, Name, Arguments)
            Else
                RunProcedure = RunApplication(ReturnVariable, Name, NoArguments)
            End If
            If RunProcedure = True Then Exit Function
        End If

        Dim VariableModel As ArrayValueModel
        Dim VariableModelIndex As Long
        VariableModelIndex = VariableIndex(VariableModel, Name)
        If Not VariableModel Is Nothing Then
            If ArraySize(Arguments) <> -1 Then
                RunProcedure = RunVariable(ReturnVariable, VariableModel, VariableModelIndex, Arguments, True)
            Else
                RunProcedure = RunVariable(ReturnVariable, VariableModel, VariableModelIndex, NoArguments,True)
            End If
        End If
    End Function

    Private Function HandleArguments(Line As String) As Variant()
        Dim Arguments() As Variant
        Arguments = GetProcedureArguments(Line)
        If ArraySize(Arguments) > -1 Then Call RecursiveInterpretArgument(Arguments)
        HandleArguments = Arguments
    End Function

    Private Function RunVariable(ReturnVariable As Variant, VariableModel As ArrayValueModel, VariableModelIndex As Long, Arguments() As Variant, Optional ReturnValue As Boolean = False) As Boolean
        Dim Procedure As ConsoleProcedure
        Set Procedure = VariableModel.Element(VariableModelIndex)

        If Returnvalue Then
            Select Case True
                Case Procedure.ReturnType Like "*As stdLambda"      : RunVariable = RunLambda(ReturnVariable, Procedure.Value, Arguments)
                Case Procedure.ReturnType Like "*As ConsoleScript"  : RunVariable = RunScript(ReturnVariable, Procedure, Arguments)
                Case Procedure.ReturnType Like "*Procedure *"       : RunVariable = RunApplication(ReturnVariable, Procedure.Name, Arguments)
                Case Procedure.ReturnType Like "*As *"              : RunVariable = GetVariableByType(ReturnVariable, Procedure.Value, Procedure.ReturnType)
                Case Else                                           : Call Errors.Add("Could not return Value")
            End Select
        Else
            Set ReturnVariable = Procedure
            ReturnVariable = True
        End If
    End Function

    Private Function RunApplication(ReturnVariable As Variant, Name As String, Arguments() As Variant) As Boolean
        On Error GoTo Error
        If IsMissing(Arguments) Then
            ReturnVariable = Application.Run(Name)
        Else
            Select Case ArraySize(Arguments)
                Case 00:   ReturnVariable = Application.Run(Name, Arguments(0))
                Case 01:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1))
                Case 02:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2))
                Case 03:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3))
                Case 04:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4))
                Case 05:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5))
                Case 06:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6))
                Case 07:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7))
                Case 08:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8))
                Case 09:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9))
                Case 10:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10))
                Case 11:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11))
                Case 12:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12))
                Case 13:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13))
                Case 14:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14))
                Case 15:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15))
                Case 16:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16))
                Case 17:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17))
                Case 18:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18))
                Case 19:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19))
                Case 20:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20))
                Case 21:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21))
                Case 22:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22))
                Case 23:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23))
                Case 24:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24))
                Case 25:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25))
                Case 26:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26))
                Case 27:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27))
                Case 28:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28))
                Case 29:   ReturnVariable = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28), Arguments(29))
                Case Else
                Call Errors.Add("Too many Arguments: " & ArraySize(Arguments))
            End Select
        End If
        If IsError(ReturnVariable) Then
            GoTo Error:
        Else
            RunApplication = True
            Exit Function
        End If
        Error:
        Call Errors.Add("Could not run Procedure. Procedure might not exist: " & Name)
    End Function

    Private Function RunLambda(ReturnVariable As Variant, Lambda As stdLambda, Arguments() As Variant) As Boolean
        On Error GoTo Error
        If ArraySize(Arguments) = -1 Then
            ReturnVariable = Lambda.Run()
        Else
            ReturnVariable = Lambda.RunEx(Arguments)
        End If
        RunLambda = True
        Exit Function
        Error:
        Call Errors.Add("could not run lambda successfully")
    End Function

    Private Function RunScript(ReturnVariable As Variant, Script As ConsoleProcedure, Optional UserArguments As Variant) As Boolean
        Dim i As Long
        Dim ScriptArguments() As String
        Dim DimVariables As ArrayValueModel
        Set DimVariables = New ArrayValueModel

        If Script.Arguments <> "No Arguments" Then
            ScriptArguments = Split(Script.Arguments, ArgSeperator)
            For i = 0 To ArraySize(ScriptArguments)
                ScriptArguments(i) = "DIM " & ScriptArguments(i)
            Next i
            Call RepeatRun(ReturnVariable, ScriptArguments, DimVariables)

            ' Assign UserArguments to the created DIM Arguments
            Dim VariableModel As ArrayValueModel
            Dim VariableModelIndex As Long
            If ArraySize(ScriptArguments) = ArraySize(UserArguments) And ArraySize(UserArguments) <> -1 Then
                For i = 0 To ArraySize(UserArguments)
                    VariableModelIndex = VariableIndex(VariableModel, ScriptArguments(i), DimVariables)
                    If Not VariableModel Is Nothing Then
                        VariableModel.Element(VariableModelIndex).Value = UserArguments(i)
                    End If
                Next
            End If
        End If

        Dim Lines() As String
        Lines = Split(Script.Value, LineSeperator)
        RunScript = RepeatRun(ReturnVariable, Lines)
        If RunScript = False Then Call Errors.Add("Script did not run successfully: " & Script.Name)
    End Function

    Private Function RunLoop(ReturnVariable As Variant, BreakCondition As String, Lines() As String) As Boolean
        Dim DimVariables As ArrayValueModel
        Dim NoError As Boolean
        Set DimVariables = New ArrayValueModel

        NoError = True
        If HandleReturnOperator(ReturnVariable, BreakCondition) Then
            Do
                If HandleReturnOperator(ReturnVariable, BreakCondition) = False Then Exit Do
                If ReturnVariable = True Then Exit Do
                RunLoop = RepeatRun(ReturnVariable, Lines, DimVariables)
            Loop
        End If
        RunLoop = NoError
    End Function

    Private Function RunCondition(ReturnVariable As Variant, Compares() As String, Lines() As String, Pointers() As Long) As Boolean
        Dim i As Long, j As Long
        Dim NoError As Boolean
        Dim DimVariables As ArrayValueModel
        Set DimVariables = New ArrayValueModel

        For i = 0 To ArraySize(Compares)
            If HandleReturnOperator(ReturnVariable, Compares(i))  = False Then Exit Function
            If ReturnVariable = True Then Exit For
        Next i
        If i > ArraySize(Compares) Then
            For j = Pointers(i) To (Pointers(i + 1) - 1)
                If NoError Then
                    NoError = Run(ReturnVariable, Lines(j), DimVariables)
                Else
                    Call      Run(ReturnVariable, Lines(j), DimVariables)
                End If
            Next j
        Else
            For j = Pointers(i) To ArraySize(Lines)
                If NoError Then
                    NoError = Run(ReturnVariable, Lines(j), DimVariables)
                Else
                    Call      Run(ReturnVariable, Lines(j), DimVariables)
                End If
            Next j
        End If
        RunCondition = NoError
    End Function

    Private Function HandleSpecial(ReturnVariable As Variant, Line As String, Optional DimVariables As ArrayValueModel) As Boolean
        Dim Words() As String
        Dim Lines() As String
        Dim Compares() As String
        Dim Pointers() As Long
        Dim MergedLines() As String
        Dim i As Long

        Select Case True
            Case UCase(Line) Like "FOR(*)", UCase(Line) Like "UNTIL(*)", UCase(Line) Like "WHILE(*)"
                Words = Split(GetParanthesesText(Line), ArgSeperator)
                If UCase(Line) Like "FOR(*)" Then
                    Lines = ArrayConvertString(ArraySplice(Words, 0, 1))
                    Call Run(ReturnVariable, Words(0))
                    HandleSpecial = RunLoop(ReturnVariable, Words(1), Lines)
                Else
                    Lines = ArrayConvertString(ArraySplice(Words, 0, 0))
                    HandleSpecial = RunLoop(ReturnVariable, Words(0), Lines)
                End If
            Case UCase(Line) Like "IF(*)", UCase(Line) Like "SELECT(*)"
                If UCase(Line) Like "IF(*)" Then
                    Words = Split(GetParanthesesText(Line), " ELSE ")
                Else
                    Words = Split(GetParanthesesText(Line), " CASE ")
                End If
                ReDim Compares(ArraySize(Words) - 1)
                ReDim Pointers(ArraySize(Words))
                For i = 0 To ArraySize(Pointers)
                    Lines = Split(Words(i), ArgSeperator)
                    Pointers(i) = ArraySize(MergedLines) + 1
                    If i =< ArraySize(Compares) Then
                        Compares(i) = Lines(0)
                        MergedLines = ArrayConvertString(ArrayMerge(MergedLines, ArrayShift(Lines)))
                    Else
                        MergedLines = ArrayConvertString(ArrayMerge(MergedLines, Lines))
                    End If
                Next i
                HandleSpecial = RunCondition(ReturnVariable, Compares, MergedLines, Pointers)
            Case UCase(Line) Like "DIM * AS *", UCase(Line) Like "PUBLIC * AS *", UCase(Line) Like "PRIVATE * AS *"
                HandleSpecial = AddNewVariable(ReturnVariable, Line, DimVariables)
            Case UCase(Line) Like "ERROR(*)"
                ReturnVariable = Errors.Element(CLng(GetParanthesesText(Line)))
                HandleSpecial = True
            Case Else
                Exit Function
        End Select
        HandleSpecial = True
    End Function

    Private Sub RecursiveInterpretArgument(ByRef Arguments() As Variant)
        Dim i As Long
        Dim CurrentArgument As Variant

        If ArraySize(Arguments) = -1 Then Exit Sub
        For i = 0 To ArraySize(Arguments)
            CurrentArgument = CStr(Arguments(i))
            Select Case True
                Case ArraySize(InStrAll(CStr(CurrentArgument), Chr(34))) >= 1
                    CurrentArgument = GetStringString(CStr(CurrentArgument))
                Case IsNumeric(CurrentArgument) Or IsDate(CurrentArgument)
                    Call InterpretVariable(CurrentArgument, CurrentArgument)
                Case ArraySize(HasOperators(CStr(CurrentArgument)))
                    Call HandleReturnOperator(CurrentArgument, CStr(CurrentArgument))
                Case Else
                    Call RunProcedure(CurrentArgument, GetProcedureName(CStr(CurrentArgument)), HandleArguments(CStr(CurrentArgument)))
            End Select
            Arguments(i) = CurrentArgument
        Next i
    End Sub

    Private Function HandleReturnOperator(ReturnVariable As Variant, Line As String) As Boolean
        Dim Values() As Variant
        Dim i As Long
        Dim PassValue As Variant
        Dim ReturnValue As Variant

        ReDim Values(0)
        Values(0) = Line
        Values = GetAllOperators(Values)

        If InterpretVariable(ReturnValue, Values(0)) = False Then Exit Function
        For i = 1 To ArraySize(Values) Step +2
            If InterpretVariable(PassValue, Values(i + 1)) = False Then Exit Function
            Select Case UCase(Values(i))
                Case "==", " IS "        : ReturnValue = (ReturnValue =   PassValue)
                Case "<>", "!=", " NOT " : ReturnValue = (ReturnValue <>  PassValue)
                Case "||", " OR "        : ReturnValue = (ReturnValue Or  PassValue)
                Case "&&", " AND "       : ReturnValue = (ReturnValue And PassValue)
                Case "", "", " XOR "     : ReturnValue = (ReturnValue Xor PassValue)
                Case "<"                 : ReturnValue = (ReturnValue <   PassValue)
                Case ">"                 : ReturnValue = (ReturnValue >   PassValue)
                Case "=<", "<="          : ReturnValue = (ReturnValue =<  PassValue)
                Case ">=", "=>"          : ReturnValue = (ReturnValue >=  PassValue)
                Case "++", "+="          : Call Run(ReturnValue, Values(i - 1) & AsgOperator & Values(i - 1) & "+" & Values(i + 1))
                Case "--", "-="          : Call Run(ReturnValue, Values(i - 1) & AsgOperator & Values(i - 1) & "-" & Values(i + 1))
                Case "**", "*="          : Call Run(ReturnValue, Values(i - 1) & AsgOperator & Values(i - 1) & "*" & Values(i + 1))
                Case "//", "/="          : Call Run(ReturnValue, Values(i - 1) & AsgOperator & Values(i - 1) & "/" & Values(i + 1))
                Case "^^", "^="          : Call Run(ReturnValue, Values(i - 1) & AsgOperator & Values(i - 1) & "^" & Values(i + 1))
                Case Else
                    If HandleCalcOperator(ReturnValue, ReturnValue, Values(i), PassValue) = False Then
                        Call Errors.Add("No valid operator")
                        Exit Function
                    End If
            End Select
        Next i
        If ArraySize(Values) = -1 Then
            Call Errors.Add("Could not handle operator: " & Line)
            Exit Function
        End If
        ReturnVariable = ReturnValue
        HandleReturnOperator = True
    End Function

    Private Function HandleCalcOperator(ReturnVariable As Variant, Value1 As Variant, Operator As Variant, Value2 As Variant) As Boolean
        On Error GoTo Error
        Select Case UCase(Operator)
            Case "+" : ReturnVariable = CDbl(Value1) + CDbl(Value2)
            Case "-" : ReturnVariable = CDbl(Value1) - CDbl(Value2)
            Case "*" : ReturnVariable = CDbl(Value1) * CDbl(Value2)
            Case "/" : ReturnVariable = CDbl(Value1) / CDbl(Value2)
            Case "^" : ReturnVariable = CDbl(Value1) ^ CDbl(Value2)
            Case "&" : ReturnVariable = CStr(Value1) & CStr(Value2)
            Case Else
                Exit Function
        End Select
        HandleCalcOperator = True
        Error:
    End Function

    Private Function InterpretVariable(ReturnVariable As Variant, Value As Variant) As Boolean
        Dim Arr() As Variant
        Dim VariablePos As Long
        Dim CurrentModel As ArrayValueModel

        VariablePos = VariableIndex(CurrentModel, GetProcedureName(CStr(Value)))
        If Not CurrentModel Is Nothing Then
            Call RunVariable(ReturnVariable, CurrentModel, VariablePos, NoArguments, True)
        Else
            Select Case True
                Case IsNumeric(Value)
                    If CLng(Value) = Round(CLng(Value), 0) Then
                        ReturnVariable = CLng(Value)
                    Else 
                        ReturnVariable = CDbl(Value)
                    End If
                Case IsDate(Value)                      : ReturnVariable = Cdate(Value)
                Case VarType(Value) = vbEmpty           : ReturnVariable = Empty
                Case VarType(Value) = vbNull            : ReturnVariable = Null
                Case VarType(Value) = vbInteger         : ReturnVariable = CInt(Value)
                Case VarType(Value) = vbLong            : ReturnVariable = CLng(Value)
                Case VarType(Value) = vbSingle          : ReturnVariable = CSng(Value)
                Case VarType(Value) = vbDouble          : ReturnVariable = CDbl(Value)
                Case VarType(Value) = vbCurrency        : ReturnVariable = CCur(Value)
                Case VarType(Value) = vbString          : ReturnVariable = CStr(Value)
                Case VarType(Value) = vbObject          : ReturnVariable = "Object"
                Case VarType(Value) = vbError           : ReturnVariable = "Error"
                Case VarType(Value) = vbBoolean         : ReturnVariable = CBool(Value)
                Case VarType(Value) = vbVariant         : ReturnVariable = CVar(Value)
                Case VarType(Value) = vbDataObject      : ReturnVariable = "DataObject"
                Case VarType(Value) = vbDecimal         : ReturnVariable = CDec(Value)
                Case VarType(Value) = vbByte            : ReturnVariable = CByte(Value)
                Case VarType(Value) = vbLongLong        : ReturnVariable = CLngLng(Value)
                Case VarType(Value) = vbUserDefinedType : ReturnVariable = "User defined Type"
                Case VarType(Value) = vbArray           : ReturnVariable = Arr
                Case Else
                    Exit Function
            End Select
        End If
        InterpretVariable = True
    End Function

    Private Function GetVariableByType(ReturnVariable As Variant, Value As Variant, DataType As String) As Boolean
        Dim Arr() As Variant
        On Error GoTo Error
        Select Case True
            Case UCase(DataType) Like "* AS EMPTY"      : ReturnVariable = Empty
            Case UCase(DataType) Like "* AS NULL"       : ReturnVariable = Null
            Case UCase(DataType) Like "* AS INTEGER"    : ReturnVariable = CInt(Value)
            Case UCase(DataType) Like "* AS LONG"       : ReturnVariable = CLng(Value)
            Case UCase(DataType) Like "* AS SINGLE"     : ReturnVariable = CSng(Value)
            Case UCase(DataType) Like "* AS DOUBLE"     : ReturnVariable = CDbl(Value)
            Case UCase(DataType) Like "* AS CURRENCY"   : ReturnVariable = Ccur(Value)
            Case UCase(DataType) Like "* AS DATE"       : ReturnVariable = Cdate(Value)
            Case UCase(DataType) Like "* AS STRING"     : ReturnVariable = CStr(Value)
            Case UCase(DataType) Like "* AS OBJECT"     : ReturnVariable = "Object"
            Case UCase(DataType) Like "* AS ERROR"      : ReturnVariable = "Error"
            Case UCase(DataType) Like "* AS BOOLEAN"    : ReturnVariable = CBool(Value)
            Case UCase(DataType) Like "* AS VARIANT"    : ReturnVariable = CVar(Value)
            Case UCase(DataType) Like "* AS DATAOBJECT" : ReturnVariable = "DataObject"
            Case UCase(DataType) Like "* AS DECIMAL"    : ReturnVariable = CDec(Value)
            Case UCase(DataType) Like "* AS BYTE"       : ReturnVariable = CByte(Value)
            Case UCase(DataType) Like "* AS LONGLONG"   : ReturnVariable = CLngLng(Value)
            Case UCase(DataType) Like "* AS UDT"        : ReturnVariable = "User defined Type"
            Case UCase(DataType) Like "* AS ARRAY"      : ReturnVariable = Arr
            Case Else
                Exit Function
        End Select
        GetVariableByType = True
        Exit Function

        Error:
    End Function

    Private Function AddNewVariable(ReturnVariable As Variant, Line As String, Optional DimVariables As ArrayValueModel) As Long
        Dim Words() As String
        Dim Pos As Long
        Dim Temp As ConsoleProcedure

        Words = Split(Line, " ")
        Set Temp = ConsoleProcedure.Create(Words(1), Words(0) & " " &  Words(2) & " " &  Words(3), "No Arguments", Empty)
        If UCase(Words(0)) Like "DIM" Then
            AddNewVariable = DimVariables.Add(Temp)
            Set ReturnVariable = DimVariables
        ElseIf UCase(Words(0)) Like "PUBLIC" Then
            AddNewVariable = PublicVariables.Add(Temp)
            Set ReturnVariable = PublicVariables
        Else
            AddNewVariable = PrivateVariables.Add(Temp)
            Set ReturnVariable = PrivateVariables
        End If
    End Function
'

' Intellisense

    Private Sub GetAllProcedures()

        Dim WB As Workbook
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        
        Dim CurrentRow As String
        Dim StartPoint As Long
        Dim EndPoint As Long
        Dim AsProcedure As String
        Dim CurrentComponent As Long

        Dim CurrentProcedurePos As Long

        Dim i As Integer
        Dim j As Integer
        Dim k As Long
        Dim Temp As Long

        Dim CurrentProcedure As ConsoleProcedure

        ' This is to get the second last space, which indicates, that its the startpoint for the returntype
        Dim TempArray() As String
        k = -1

        For Each WB In Workbooks
            Set VBProj = WB.VBProject
            k = k + 1
            For Each VBComp In VBProj.VBComponents

                Set CurrentProcedure = New ConsoleProcedure
                CurrentProcedure.Name       = VBComp.Name
                CurrentProcedure.Arguments  = "No Arguments"
                CurrentProcedure.Value      = "No Value"
                CurrentProcedure.ReturnType = IIf(VBComp.Type = vbext_ComponentType.vbext_ct_StdModule, "As Module", _
                                              IIf(VBComp.Type = vbext_ComponentType.vbext_ct_ClassModule, "As Class", "As Form"))
                CurrentComponent = Publics(k).Add(0, CurrentProcedure)

                Set CodeMod = VBComp.CodeModule
                For i = 1 To CodeMod.CountOfLines
                    Set CurrentProcedure = New ConsoleProcedure
                    CurrentRow = CodeMod.Lines(i, 1)
                    AsProcedure = ""
                    If UCase(CurrentRow) Like "*PUBLIC *" And InStr(1, CurrentRow, "'") = 0 And Not UCase(CurrentRow) Like "*" & Chr(34) & "*PUBLIC*" & Chr(34) & "*" Then
                        If (UCase(CurrentRow) Like "* FUNCTION *" Or UCase(CurrentRow) Like "* SUB *" Or UCase(CurrentRow) Like "* SET *" Or UCase(CurrentRow) Like "* LET *" Or UCase(CurrentRow) Like "* GET *" Or UCase(CurrentRow) Like "* EVENT *") Then
                            ' A Procedure
                            '                          |----------|
                            '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                            AsProcedure = "Public Procedure "
                            Select Case True
                                Case UCase(CurrentRow) Like "*PUBLIC STATIC SUB *(*)*"      : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC STATIC SUB ")      + Len("PUBLIC STATIC SUB ")      : EndPoint = InStr(1, UCase(CurrentRow), "("): CurrentProcedure.Name = MidP(CurrentRow, StartPoint, EndPoint - 1)
                                Case UCase(CurrentRow) Like "*PUBLIC SUB *(*)*"             : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC SUB ")             + Len("PUBLIC SUB ")             : EndPoint = InStr(1, UCase(CurrentRow), "("): CurrentProcedure.Name = MidP(CurrentRow, StartPoint, EndPoint - 1)
                                Case UCase(CurrentRow) Like "*PUBLIC STATIC FUNCTION *(*)*" : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC STATIC FUNCTION ") + Len("PUBLIC STATIC FUNCTION ") : EndPoint = InStr(1, UCase(CurrentRow), "("): CurrentProcedure.Name = MidP(CurrentRow, StartPoint, EndPoint - 1)
                                Case UCase(CurrentRow) Like "*PUBLIC FUNCTION *(*)*"        : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC FUNCTION ")        + Len("PUBLIC FUNCTION ")        : EndPoint = InStr(1, UCase(CurrentRow), "("): CurrentProcedure.Name = MidP(CurrentRow, StartPoint, EndPoint - 1)
                                Case UCase(CurrentRow) Like "*PUBLIC EVENT *(*)*"           : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC EVENT ")           + Len("PUBLIC EVENT ")           : EndPoint = InStr(1, UCase(CurrentRow), "("): CurrentProcedure.Name = MidP(CurrentRow, StartPoint, EndPoint - 1)
                                Case UCase(CurrentRow) Like "*PUBLIC PROPERTY GET *(*)*"    : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC PROPERTY GET ")    + Len("PUBLIC PROPERTY GET ")    : EndPoint = InStr(1, UCase(CurrentRow), "("): CurrentProcedure.Name = MidP(CurrentRow, StartPoint, EndPoint - 1)
                                Case UCase(CurrentRow) Like "*PUBLIC PROPERTY SET *(*)*"    : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC PROPERTY SET ")    + Len("PUBLIC PROPERTY SET ")    : EndPoint = InStr(1, UCase(CurrentRow), "("): CurrentProcedure.Name = MidP(CurrentRow, StartPoint, EndPoint - 1)
                                Case UCase(CurrentRow) Like "*PUBLIC PROPERTY LET *(*)*"    : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC PROPERTY LET ")    + Len("PUBLIC PROPERTY LET ")    : EndPoint = InStr(1, UCase(CurrentRow), "("): CurrentProcedure.Name = MidP(CurrentRow, StartPoint, EndPoint - 1)
                                Case Else
                            End Select
                            '                                       |------------------------------|
                            '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                                CurrentProcedure.Arguments = GetParanthesesText(CurrentRow)
                                If CurrentProcedure.Arguments = CurrentRow Then CurrentProcedure.Arguments = "No Arguments"
                        Else
                            ' A Variable
                            '          |----------|
                            '   Public VariableName As Variant
                            Select Case True
                                Case UCase(CurrentRow) Like "*PUBLIC CONST *": StartPoint = InStr(1, UCase(CurrentRow), "*PUBLIC CONST *") + Len("PUBLIC CONST "): EndPoint = InStr(1, UCase(CurrentRow), " AS "): CurrentProcedure.Name = MidP(CurrentRow, StartPoint + 1, EndPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC ENUM *":  StartPoint = InStr(1, UCase(CurrentRow), "*PUBLIC ENUM *")  + Len("PUBLIC ENUM "):  EndPoint = Len(CurrentRow)                    : CurrentProcedure.Name = MidP(CurrentRow, StartPoint + 1, EndPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC TYPE *":  StartPoint = InStr(1, UCase(CurrentRow), "*PUBLIC TYPE *")  + Len("PUBLIC TYPE "):  EndPoint = Len(CurrentRow)                    : CurrentProcedure.Name = MidP(CurrentRow, StartPoint + 1, EndPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC *":       StartPoint = InStr(1, UCase(CurrentRow), "*PUBLIC *")       + Len("PUBLIC "):       EndPoint = InStr(1, UCase(CurrentRow), " AS "): CurrentProcedure.Name = MidP(CurrentRow, StartPoint + 1, EndPoint)
                                Case Else
                            End Select
                        End If
                        '                                                                        |---------|
                        '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                        TempArray = Split(CurrentRow, " ")
                        CurrentProcedure.ReturnType = AsProcedure & TempArray(ArraySize(TempArray) - 1) & " " & TempArray(ArraySize(TempArray))
                        ' If last character is ")", then it returns void
                        If Mid(TempArray(ArraySize(TempArray)), Len(TempArray(ArraySize(TempArray))), 1) = ")" Then
                            CurrentProcedure.ReturnType = AsProcedure & "As Void"
                        End If
                        CurrentProcedure.Value = "No Value"
                        Call Publics(k).Add(CurrentComponent, CurrentProcedure)
                    End If
                    NoLines:
                Next
            Next
        Next
                
    End Sub

    Public Function GetIntellisenseList(Text As String) As ArrayValueModel

        Dim Words() As String
        Dim CurrentWorkbook As Long
        Dim i As Long, j As Long, k As Long
        Dim NameSpaces As New ArrayValueModel
        Dim CurrentComponent As Long
        Dim Index As Long

        Dim ReturnArray As New ArrayValueModel

        Words = Split(Text, ".")
        If ArraySize(Words) = -1 Then Exit Function
        

        Select Case True
            Case ArraySize(Words) = 0
                For i = 0 To ArraySize(Publics) ' Workbooks
                    Call NameSpaces.Add(Publics(i).Value(0))
                Next i

                For i = 0 To ArraySize(Publics) ' For each workbook
                    For j = 0 To ArraySize(Publics(i).Branches(0)) ' For each Component
                        CurrentWorkbook = Publics(i).Branches(0)(j)
                        If Publics(i).Value(CurrentWorkbook).ReturnType = "As Module" Then ' If Component is module
                            For k = 0 To ArraySize(Publics(i).Branches(CurrentWorkbook)) ' For each Public in module
                                CurrentComponent = Publics(i).Branches(CurrentWorkbook)(k)
                                Call NameSpaces.Add(Publics(i).Value(CurrentComponent))
                            Next k
                        End If
                    Next j
                Next i

                For i = 0 To ArraySize(PublicVariables.Arr)
                    Call NameSpaces.Add(PublicVariables.Element(i))
                Next i
                For i = 0 To ArraySize(PrivateVariables.Arr)
                    Call NameSpaces.Add(PrivateVariables.Element(i))
                Next i

            Case ArraySize(Words) = 1
                For i = 0 To ArraySize(Publics)
                    If Words(0) = CStr(Publics(i).Value(0).Name) Then Exit For
                Next i
                If i > ArraySize(Publics) Then ' If Not a Workbook
                    Exit Function
                Else
                    For j = 0 To ArraySize(Publics(i).Branches(0)) ' For Each Component
                        Index = Publics(i).Branches(0)(j)
                        Call NameSpaces.Add(Publics(i).Value(Index))
                    Next j
                End If

            Case ArraySize(Words) = 2
                For i = 0 To ArraySize(Publics)
                    If Words(0) = CStr(Publics(i).Value(0).Name) Then Exit For
                Next i
                If i > ArraySize(Publics) Then ' If not a Component
                    Exit Function
                Else
                    For j = 0 To ArraySize(Publics(i).Branches(0))
                        Index = Publics(i).Branches(0)(j)
                        If Publics(i).Value(Index).Name = Words(1) Then Exit For
                    Next j
                    If j > ArraySize(Publics(i).Branches(0)) Then
                        Exit Function
                    Else
                        For k = 0 To ArraySize(Publics(i).Branches(j))
                            CurrentComponent = Publics(i).Branches(j)(k)
                            Call NameSpaces.Add(Publics(i).Value(CurrentComponent))
                        Next k
                    End If
                End If

            Case Else
        End Select

        For i = 0 To ArraySize(NameSpaces.Arr)
            If ArraySize(Words) = -1 Then
                Call ReturnArray.Add(NameSpaces.Element(i))
            Else
                If CStr(NameSpaces.Element(i).Name) Like Words(ArraySize(Words)) & "*" Then
                    Call ReturnArray.Add(NameSpaces.Element(i))
                End If
            End If
        Next i

        Set GetIntellisenseList = ReturnArray
    End Function
'
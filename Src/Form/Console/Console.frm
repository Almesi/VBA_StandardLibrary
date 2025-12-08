VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Console 
   Caption         =   "Console"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "Console.frx":0000
   ShowModal       =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IConsoleView

' Private Variables
    
    Private       CurrentLineIndex      As Long
    Private Const Recognizer            As String = "\>>>"
    Private Const ArgSeperator          As String = ", " 
    Private Const AsgOperator           As String = " = "
    Private Const LineSeperator         As String = "LINEBREAK/()\"
    Private       IsCancelled           As Boolean

    Private PreviousCommands As ArrayValueModel
    Private Interpreter      As ConsoleInterpreter
    Private ColorDef         As ColorDefinition
    Private ColorRuleSet     As TextColoring

    Private WorkMode As Long
    Private Enum WorkModeEnum
        Logging = 0
        UserInputt = 1
        Multiline = 2
        Script = 3
        Idle = 4
    End Enum

'

' Public Console Functions

    Private Function IConsoleView_Run(Model As Object) As Boolean
        Interpreter.PublicVar = Model
        Call Show(vbModal)
        Set Model = Interpreter.PublicVar
        IConsoleView_Run = Not IsCancelled
    End Function

    Public Sub IConsoleView_Load(Optional n_ColorDef As ColorDefinition = Nothing, Optional ScreenHeight As Long = 5000, Optional ScreenWidth As Long = 3000)
        If ColorDef Is Nothing Then
            Set ColorDef = New ColorDefinition
            Call AssignColor()
            Set ColorRuleSet = New TextColoring
            ColorRuleSet.ColorDef = ColorDef
            Call AssignRules()
        Else
            ColorDef = n_ColorDef
        End If

        
        ConsoleText.Text = GetStartText & PrintStarter
        ConsoleText.SelStart = 0
        ConsoleText.SelLength = Len(ConsoleText.Text)
        ConsoleText.SelColor = ColorDef.Color("System")
        ConsoleText.SelStart = Len(ConsoleText.Text)

        Call SetUpNewLine()
        ScrollHeight = ScreenHeight
        ScrollWidth  = ScreenWidth
    End Sub

    Public Function IConsoleView_CheckPredeclaredAnswer(ReturnVariable As Variant, Message As Variant, AllowedValues As ArrayValueModel) As Long
        Dim i As Long
        Dim Found As Boolean

        Message = Message & "("
        For i = 0 To ArraySize(AllowedValues.Arr)
            Message = Message & AllowedValues.Element(i) & "|"
        Next i
        Message = Message & ") "
        Call IConsoleView_PrintConsole(Message, NoColor, NoColorLength)

        WorkMode = WorkModeEnum.UserInputt
        Do While WorkMode = WorkModeEnum.UserInputt
            DoEvents
            If ReturnVariable <> Empty Then
                ReturnVariable = Replace(ReturnVariable, Message, "")
                For i = 0 To ArraySize(AllowedValues.Arr)
                    If AllowedValues.Element(i) = ReturnVariable Then
                        IConsoleView_CheckPredeclaredAnswer = i
                        Found = True
                        Exit For
                    End If
                Next i
                If Found = False Then Call IConsoleView_PrintConsole(Message, NoColor, NoColorLength)
            End If
        Loop
        WorkMode = WorkModeEnum.Logging
        Call IConsoleView_PrintConsole(PrintStarter, NoColor, NoColorLength)
    End Function

    Public Sub IConsoleView_PrintEnter(Text As String, Colors() As Long, ColorLength() As Long)
        Call IConsoleView_PrintConsole(Text & vbCrLf, Colors, ColorLength)
    End Sub

    Public Sub IConsoleView_PrintConsole(Text As String, Colors() As Long, ColorLength() As Long)
        
        Dim i As Long
        Dim StartPoint As Long
        Dim Offset As Long

        If ArraySize(Colors) = -1 Then
            ReDim Colors(0)
            Colors(0) = ColorDef.Color("System")
        End If
        If ArraySize(ColorLength) = -1 Then
            ReDim ColorLength(0)
            ColorLength(0) = Len(Text)
        End If

        If ArraySize(Colors) <> ArraySize(ColorLength) Then Exit Sub

        StartPoint = Len(ConsoleText.Text)
        Offset = 1
        For i = 0 To ArraySize(ColorLength)
            ConsoleText.SelStart = StartPoint
            ConsoleText.SelLength = 0
            ConsoleText.SelColor = Colors(i)
            ConsoleText.SelText = Mid(Text, Offset, ColorLength(i))
            Offset = Offset + ColorLength(i)
            StartPoint = StartPoint + ColorLength(i)
        Next
        Call SetUpNewLine
        ConsoleText.SelStart = StartPoint

    End Sub

    Public Sub IConsoleView_ColorConsole(StartPoint As Long, Colors() As Long, ColorLength() As Long)
        Dim i As Long
        Dim Offset As Long

        If ArraySize(Colors) = -1 Then
            ReDim Colors(0)
            Colors(0) = ColorDef.Color("System")
        End If
        If ArraySize(ColorLength) = -1 Then
            ReDim ColorLength(0)
            ColorLength(0) = Len(ConsoleText.Text)
        End If
        If ArraySize(Colors) <> ArraySize(ColorLength) Then Exit Sub

        Offset = StartPoint
        For i = 0 To ArraySize(ColorLength)
            ConsoleText.SelStart = Offset
            ConsoleText.SelLength = ColorLength(i)
            ConsoleText.SelColor = Colors(i)
            Offset = Offset + ColorLength(i)
        Next
        ConsoleText.SelStart = Len(ConsoleText.Text)
    End Sub
'

' Initialization

    Private Sub UserForm_Initialize()
        Set PreviousCommands = New ArrayValueModel
        Set Interpreter = New ConsoleInterpreter
    End Sub

    Private Sub InitializeTree()
        Dim i As Long
        Dim WB As Workbook
        ReDim Trees(Workbooks.Count - 1)
        i = 0
        For Each WB In Workbooks
            Set Trees(i) = std_Tree.Create()
            Trees(i).Value(0) = ConsoleProcedure.Create(WB.VBProject.Name, "As VBProject", "No Arguments", "No Value")
            i = i + 1
        Next
    End Sub

    Private Sub UserForm_Terminate()
    End Sub

    Private Sub OnCancel()
        IsCancelled = True
        Call Hide()
    End Sub

    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        If CloseMode = VbQueryClose.vbFormControlMenu Then
            Cancel = True
            OnCancel
        End If
    End Sub
'

' Get/Set Values

    Private Function GetMaxSelStart() As Long
        Dim Temp(1) As Long
        Temp(0) = ConsoleText.SelStart
        Temp(1) = ConsoleText.SelLength
        ConsoleText.SelStart = Len(ConsoleText.Text)
        GetMaxSelStart = ConsoleText.SelStart
        ConsoleText.SelStart = Temp(0)
        ConsoleText.SelLength = Temp(1)
    End Function

    Private Function PrintStarter() As String
        PrintStarter = ThisWorkbook.Path & Recognizer
    End Function

    Private Function GetStartText() As String
        GetStartText =                   _
        "VBA Console [Version 2.0]" & vbCrLf & _
        "No Rights reserved"        & vbCrLf & _
        vbCrLf
    End Function

    Private Function GetTextLength(Text As String, Seperator As String, Optional IndexBreakPoint As Long = -2) As Long
        Dim i As Long
        Dim Lines() As String
        Lines = Split(Text, Seperator)
        If IndexBreakPoint = -2 Then IndexBreakPoint = ArraySize(Lines)
        For i = 0 To IndexBreakPoint
            GetTextLength = GetTextLength + Len(Lines(i)) + 1
        Next i
    End Function

    Private Function GetLine(Text As String, Index As Long) As String
        Dim Lines() As String
        Dim SearchString As String
        Dim Point As Long

        Lines = Split(Text, vbCrLf)
        If Index > 0 And Index <= ArraySize(Lines) + 1 Then
            SearchString = Lines(Index)
            Point = InStr(1, SearchString, Recognizer)
            If Point = 0 Then
                GetLine = ""
            Else
                GetLine = MidP(SearchString, Point + 1, Len(Recognizer))
            End If
        End If
    End Function

    Private Function GetWord(Text As String, Optional Index As Long = -1) As String
        Dim Words() As String
        Words = Split(Text, " ")
        If Index = -1 Then Index = ArraySize(Words)
        If ArraySize(Words) > -1 Then GetWord = Words(Index)
    End Function

    Private Sub SetUpNewLine()
        CurrentLineIndex = ArraySize(Split(ConsoleText.Text, vbCrLf))
    End Sub

    Private Function GetMultilines() As String
        Dim CurrentLine As String
        Dim CurrentCount As Long
        Dim Text As String

        CurrentCount = CurrentLineIndex - 2
        Text = GetLine(ConsoleText.Text, CurrentCount)
        Do Until UCase(GetLine(ConsoleText.Text, CurrentCount)) = "MULTILINE" Or UCase(GetLine(ConsoleText.Text, CurrentCount)) = "SCRIPT" Or CurrentCount = 0
            Text = GetLine(ConsoleText.Text, CurrentCount)
            CurrentCount = CurrentCount - 1
            CurrentLine = Text & vbCrLf & CurrentLine
        Loop
        GetMultilines = Mid(CurrentLine, 1, Len(CurrentLine))
    End Function

    Private Function WorkMultilines(Text As String) As String
        If WorkMode = WorkModeEnum.Script Then
            Dim Name As String
            Dim Arguments As String
            Dim Lines As String

            Dim TempCount As Long
            Dim TempName As String
            Dim Temp As ArrayValueModel

            Text = Replace(Text, "_" & vbCrLf, "")
            Text = Replace(Text, vbCrLf, LineSeperator)
            TempCount = InStr(1, Text, LineSeperator)
            TempName = Mid(Text, 1, TempCount - 1)
            If InStr(1, TempName, "(") = 0 Then TempName = TempName & "()"

            Name = GetProcedureName(TempName)
            Arguments = GetParanthesesText(TempName)
            Lines = Mid(Text, TempCount + Len(LineSeperator), Len(Text))

            Set Temp = Interpreter.PrivateVar
            Call Temp.Add(ConsoleProcedure.CreateScript(Name, Arguments, Lines))
            Interpreter.PrivateVar = Temp

            WorkMode = WorkModeEnum.Logging
        Else
            Text = Replace(Text, vbCrLf, "")
            WorkMode = WorkModeEnum.Logging
            WorkMultilines = Text
        End If
    End Function

    Private Function ObjData(Variable As Variant) As String
        On Error GoTo Error
        ObjData = Cstr(Variable)
        Exit Function
        Error:
        ObjData = "Not Printable"
    End Function

'

' Handle Input

    Private Sub ConsoleText_KeyDown(pKey As Long, ByVal ShiftKey As Integer)
        If WorkMode = WorkModeEnum.Idle Then pKey = 0: Exit Sub
        If pKey = 13 Then
            ConsoleText.SelStart  = GetMaxSelStart
            ConsoleText.Sellength = 0
            ConsoleText.SelColor  = ColorDef.Color("Basic")
        End If
    End Sub

    Private Sub ConsoleText_KeyUp(pKey As Long, ByVal ShiftKey As Integer)
        
        Static PreviousCommandsIndex As Long
        Dim Lines As Variant
        If WorkMode = WorkModeEnum.Idle Then Exit Sub
        Lines = Split(ConsoleText.Text, vbCrLf)
        Call SetUpNewLine()
        Select Case pKey
            Case vbKeyReturn
                Call HandleEnter()
                Call PreviousCommands.Add(GetLine(ConsoleText.Text, ArraySize(Lines) - 1))
                PreviousCommandsIndex = ArraySize(PreviousCommands) + 1
                Call SetPositions
                ConsoleText.SelStart  = GetMaxSelStart
                ConsoleText.Sellength = 0
                ConsoleText.SelColor  = ColorDef.Color("Basic")
            Case vbKeyUp, vbKeyDown
                If Workmode = WorkModeEnum.Logging Then
                    If pKey = vbKeyUp Then
                        PreviousCommandsIndex = PreviousCommandsIndex - 1
                        If PreviousCommandsIndex < 0 Then PreviousCommandsIndex = ArraySize(PreviousCommands.Arr)
                    End If
                    If pKey = vbKeyDown Then
                        PreviousCommandsIndex = PreviousCommandsIndex + 1
                        If PreviousCommandsIndex > ArraySize(PreviousCommands) Then PreviousCommandsIndex = 0
                    End If
                    ConsoleText.SelStart  = GetTextLength(ConsoleText.Text, vbCrLf, ArraySize(Lines) - 1)
                    ConsoleText.SelLength = Len(ConsoleText.Text)
                    ConsoleText.SelText   = PrintStarter & PreviousCommands.Element(PreviousCommandsIndex)
                End If
            Case Else
                Call HandleOtherKeys(pKey, ShiftKey)
        End Select

    End Sub

    Private Function HandleEnter() As Variant

        Dim i As Long
        Dim Line As String
        Dim Value As Variant

        Line = GetLine(ConsoleText.Text, CurrentLineIndex - 1)

        Repeat:
        Select Case WorkMode
            Case WorkModeEnum.Logging
                If HandleSpecial(Value, Line) Then
                    If Value = "Repeat" Then
                        GoTo Repeat
                    Else
                        Call IConsoleView_PrintEnter(ObjData(Value), NoColor, NoColorLength)
                    End If
                Else
                    Call Interpreter.Run(Value, Line)
                    Call IConsoleView_PrintEnter(ObjData(Value), NoColor, NoColorLength)
                End If
            Case WorkModeEnum.Multiline, WorkModeEnum.Script
                If UCase(Line) = "ENDSCRIPT" Or UCase(Line) = "ENDMULTILINE" Then
                    Value = WorkMultilines(GetMultilines())
                    If Value <> Empty Then GoTo Repeat
                End If
        End Select
        If Workmode = WorkModeEnum.Multiline Or WorkMode = WorkModeEnum.Script Then
        Else
            Call IConsoleView_PrintConsole(PrintStarter, NoColor, NoColorLength)
        End If
    End Function

    Private Function HandleSpecial(ReturnVariable As Variant, Line As String) As Boolean
        Select Case True
            Case UCase(Line) Like "HELP"           : ReturnVariable = HandleHelp()
            Case UCase(Line) Like "CLEAR"          : ReturnVariable = HandleClear()
            Case UCase(Line) Like "MULTILINE"      : ReturnVariable = "Repeat": Workmode = WorkModeEnum.Multiline
            Case UCase(Line) Like "SCRIPT"         : ReturnVariable = "Repeat": Workmode = WorkModeEnum.Script
            Case UCase(Line) Like "IDLE"           : ReturnVariable = "Repeat": WorkMode = WorkModeEnum.Idle
            Case UCase(Line) Like "CANCEL"
                Call HandleClear()
                Call OnCancel()
            Case UCase(Line) Like "EXIT"
                Hide
            Case UCase(Line) Like "ERROR(*)"
                If IsNumeric(GetParanthesesText(Line)) Then
                    Call GetError(CLng(GetParanthesesText(Line)))
                Else
                    Call GetError(-1)
                End If
            Case Else
                Exit Function
        End Select
        HandleSpecial = True
    End Function

    Private Function HandleClear() As String
        ConsoleText.Text = ""
        ConsoleText.SelStart = 0
        ConsoleText.SelLength = Len(ConsoleText.Text)
        ConsoleText.SelColor = ColorDef.Color("Basic")
        Call SetUpNewLine
        HandleClear = " "
    End Function

    Private Function HandleHelp() As String
        HandleHelp = "Here is the Link to the documentation: https://github.com/Almesi/VBA_StandardLibrary/tree/main/Src/Form/Console/ConsoleTutorial.md"
    End Function

    Private Function HandleOtherKeys(pKey As Long, ByVal ShiftKey As Integer) As String

        Static CapitalKey As Boolean
        Dim AsciiChar As String
        Dim CurrentWord As String
        Dim CurrentLine As String
        Dim CurrentSelection(1) As Long
        
        ' Adjust for Shift key (Uppercase letters, special characters)
        CurrentLine = GetLine(ConsoleText.Text, CurrentLineIndex)
        CurrentWord = GetWord(CurrentLine)
        If pKey = vbKeyCapital Then
            CapitalKey = CapitalKey Xor True
            GoTo SkipKey
        End If
        If CapitalKey = True Then ShiftKey = 1
        Select Case ShiftKey
            Case 0
                ' Base character
                Select Case pKey
                    Case vbKeyA To vbKeyZ:      AsciiChar = LCase(Chr(pKey))
                    Case vbKey0 To vbKey9:      AsciiChar = Chr(pKey)
                    Case vbKeySpace:            AsciiChar = " "
                    Case vbKeyBack:             AsciiChar = Chr(8) ' Backspace
                    Case vbKeyReturn:           AsciiChar = Chr(13) ' Carriage Return
                    Case vbKeyTab:              AsciiChar = Chr(9) ' Tab
                    Case vbKeyMultiply:         AsciiChar = "*"
                    Case vbKeyAdd, 187:         AsciiChar = "+"
                    Case vbKeySubtract, 189:    AsciiChar = "-"
                    Case vbKeyDecimal, 190:     AsciiChar = "."
                    Case vbKeyDivide:           AsciiChar = "/"
                    Case 188:                   AsciiChar = ","
                    Case 191:                   AsciiChar = "#"
                    Case 226:                   AsciiChar = "<"
                    Case vbKeyRight:            AsciiChar = "RIGHT"
                    Case vbKeyLeft:             AsciiChar = "LEFT"
                    Case vbKeyUp:               AsciiChar = "UP"
                    Case vbKeyDown:             AsciiChar = "DOWN"
                End Select
            Case = 1
                Select Case pKey
                    Case vbKeyA To vbKeyZ:      AsciiChar = UCase(AsciiChar)
                    Case vbKey1:                AsciiChar = "!"
                    Case vbKey2:                AsciiChar = Chr(34) ' """
                    Case vbKey3:                AsciiChar = "§"
                    Case vbKey4:                AsciiChar = "$"
                    Case vbKey5:                AsciiChar = "%"
                    Case vbKey6:                AsciiChar = "&"
                    Case vbKey7:                AsciiChar = "/"
                    Case vbKey8:                AsciiChar = "("
                    Case vbKey9:                AsciiChar = ")"
                    Case vbKey0:                AsciiChar = "="
                    Case 187:                   AsciiChar = "*"
                    Case 188:                   AsciiChar = ";"
                    Case 189:                   AsciiChar = "_"
                    Case 190:                   AsciiChar = ":"
                    Case 191:                   AsciiChar = "'"
                    Case 226:                   AsciiChar = ">"
                End Select
            Case 2

            Case 3
                    Case 226:                   AsciiChar = "|"
        End Select
        Select Case AsciiChar
            Case "RIGHT"
                If ConsoleText.SelStart = GetMaxSelStart Then
                    If IntellisenseList.ListCount > 0 Then IntellisenseList.SetFocus
                End If
            Case Else
                IntelliSenseList.Clear
                Call SetUpIntelliSenseList(CurrentWord)
        End Select
        SkipKey:
        Call SetPositions
        CurrentSelection(0) = ConsoleText.SelStart
        CurrentSelection(1) = ConsoleText.Sellength
        Call ColorLine()
        ConsoleText.SelStart = CurrentSelection(0)
        ConsoleText.SelLength = CurrentSelection(1)
        ConsoleText.SelColor = ColorDef.Color("Basic")
        HandleOtherKeys = AsciiChar

    End Function
    
'

' Coloring

    Private Sub AssignColor()
        Call ColorDef.AddColor("Operator"   , RGB(255, 000, 000))
        Call ColorDef.AddColor("Statement"  , RGB(255, 000, 255))
        Call ColorDef.AddColor("Keyword"    , RGB(000, 000, 255))
        Call ColorDef.AddColor("Parantheses", RGB(170, 170, 000))
        Call ColorDef.AddColor("Datatype"   , RGB(000, 170, 000))
        Call ColorDef.AddColor("String"     , RGB(255, 165, 000))
        Call ColorDef.AddColor("Procedure"  , RGB(255, 255, 000))
        Call ColorDef.AddColor("Script"     , RGB(255, 128, 128))
        Call ColorDef.AddColor("Lambda"     , RGB(170, 000, 170))
        Call ColorDef.AddColor("Variable"   , RGB(000, 255, 255))
        Call ColorDef.AddColor("Value"      , RGB(000, 255, 000))
        Call ColorDef.AddColor("Basic"      , RGB(255, 255, 255))
        Call ColorDef.AddColor("System"     , RGB(170, 170, 170))
    End Sub

    Private Sub AssignRules()
        Call ColorRuleSet.MakeRule("Operator"   , "$1=" & Chr(34) & "+"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "*"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "/"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "-"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "^"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & ":"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & ";"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "<"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & ">"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "="        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "!"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "|"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "?"        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & ","        & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "NOT"      & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "AND"      & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "OR"       & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "XOR"      & Chr(34))

        Call ColorRuleSet.MakeRule("Statement"  , "$1=" & Chr(34) & "IF"       & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "THEN"     & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "ELSE"     & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "END"      & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "FOR"      & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "EACH"     & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "NEXT"     & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "DO"       & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "WHILE"    & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "LOOP"     & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "SELECT"   & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "CASE"     & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "EXIT"     & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "CONTINUE" & Chr(34))

        Call ColorRuleSet.MakeRule("Keyword"    , "$1=" & Chr(34) & "IF"       & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "DIM"      & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "PUBLIC"   & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "PRIVATE"  & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "GLOBAL"   & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "TRUE"     & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "FALSE"    & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "FUNCTION" & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "SUB"      & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "REDIM"    & Chr(34) & " or " & _
                                                  "$1=" & Chr(34) & "PRESERVE" & Chr(34))

        Call ColorRuleSet.MakeRule("Parantheses", "$1="      & Chr(34) & "("                         & Chr(34) & " or $1= "     & Chr(34) & ")"    & Chr(34))
        Call ColorRuleSet.MakeRule("Datatype"   , "$1 like " & Chr(34) & "AS*"                       & Chr(34))
        Call ColorRuleSet.MakeRule("String"     , "$1 like " & Chr(34) & "Chr(34)" & "*" & "Chr(34)" & Chr(34))
        Call ColorRuleSet.MakeRule("Procedure"  , "$1 like " & Chr(34) & "* PROCEDURE *"             & Chr(34))
        Call ColorRuleSet.MakeRule("Script"     , "$1 like " & Chr(34) & "* AS CONSOLESCRIPT"        & Chr(34))
        Call ColorRuleSet.MakeRule("Lambda"     , "$1 like " & Chr(34) & "* AS STDLAMBDA"            & Chr(34))
        Call ColorRuleSet.MakeRule("Variable"   , "$1 like " & Chr(34) & "* AS *"                    & Chr(34) & " or $1 like " & Chr(34) & "AS *" & Chr(34))
        Call ColorRuleSet.MakeRule("Value"      , "isnumeric($1)")
        Call ColorRuleSet.MakeRule("Basic"      , "1=1")
        Call ColorRuleSet.MakeRule("System"     , "1=1")
    End Sub

    Private Sub ColorLine()
        Dim LineStarters() As Long
        Dim CurrentLinePoint As Long
        Dim Text As String
        Dim Colors() As Long
        Dim ColorLengths() As Long
        Dim LenToSelStartOffset As Long: LenToSelStartOffset = 1
        Dim StartSize As Long

        LineStarters = InStrAll(ConsoleText.Text, Recognizer)
        StartSize = LineStarters(ArraySize(LineStarters)) + Len(Recognizer) 
        CurrentLinePoint = StartSize - (ArraySize(InStrAll(ConsoleText.Text, vbCrLf)) + 1) - LenToSelStartOffset
        Text = MidP(ConsoleText.Text, StartSize, Len(ConsoleText.Text))
        Call ColorText(Colors, ColorLengths, Text)

        Call IConsoleView_ColorConsole(CurrentLinePoint, Colors, ColorLengths)
    End Sub

    Private Sub ColorText(Colors() As Long, ColorLengths() As Long, ByVal Text As String)
        Dim Size As Long
        Dim Words() As String
        Dim Points() As Long
        Dim i As Long
        Dim Procedure As ConsoleProcedure

        If Len(Text) =< 0 Then Exit Sub

        Call ColorString(Colors, ColorLengths, Text)
        Call ColorOperator(Colors, ColorLengths, Text)

        Size = ArraySize(Colors) + 1
        Words = Split(Text, " ")
        For i = 0 To ArraySize(Words)
            ReDim Preserve Colors(Size)
            ReDim Preserve ColorLengths(Size)
            Select Case True
                Case UCase(Words(i)) = "AS"
                    If (i + 1) > ArraySize(Words) Then
                        Colors(Size)       = ColorRuleSet.GetColor(UCase(Words(i)))
                        ColorLengths(Size) = Len(Words(i))
                    Else
                        Colors(Size)       = ColorRuleSet.GetColor(UCase(Words(i) & " " & Words(i + 1)))
                        ColorLengths(Size) = Len(Words(i)) + Len(Words(i + 1)) + 1
                        i = i + 1
                    End If
                Case Interpreter.GetProcedure(Procedure, Words(i))
                    Colors(Size)       = ColorRuleSet.GetColor(UCase(Procedure.ReturnType))
                    ColorLengths(Size) = Len(GetProcedureName(Words(i)))
                Case Else           
                    Colors(Size)       = ColorRuleSet.GetColor(UCase(Words(i)))
                    ColorLengths(Size) = Len(Words(i))
            End Select
            If i < ArraySize(Words) Then ColorLengths(Size) = ColorLengths(Size) + 1
            Size = Size + 1
        Next i
    End Sub

    Private Sub ColorString(Colors() As Long, ColorLengths() As Long, ByVal Text As String)
        Dim Points() As Long
        Dim PreviousPoint As Long
        Dim LastPoint As Long
        Dim Size As Long
        Dim i As Long

        If Len(Text) =< 0 Then Exit Sub

        Points = InStrAll(Text, Chr(34))
        PreviousPoint = 1
        If ArraySize(Points) > 0 Then ' Ensure " Character was found
            For i = 0 To ArraySize(Points) Step +2
                LastPoint = Points(i) - 1
                Call ColorText(Colors, ColorLengths, MidP(Text, PreviousPoint, LastPoint))
                Size = ArraySize(Colors) + 1

                ReDim Preserve Colors(Size)
                ReDim Preserve ColorLengths(Size)
                Colors(Size)       = ColorDef.Color("String")
                ColorLengths(Size) = Points(i + 1) - Points(i) + 1

                PreviousPoint = LastPoint
                Size = Size + 1
            Next i
            Call ColorText(Colors, ColorLengths, MidP(Text, Points(i - 1) + 1, Len(Text)))
        ElseIf ArraySize(Points) = 0 Then
            Call ColorText(Colors, ColorLengths, MidP(Text, 1, Points(0) - 1))
            Size = ArraySize(Colors) + 1
            ReDim Preserve Colors(Size)
            ReDim Preserve ColorLengths(Size)
            Colors(Size)       = ColorDef.Color("String")
            ColorLengths(Size) = Len(Text) - Points(0) + 1
        End If
    End Sub

    Private Sub ColorOperator(Colors() As Long, ColorLengths() As Long, ByVal Text As String)
        Dim SpecialChars() As Variant
        Dim Points() As Long
        Dim Counter As Long
        Dim Size As Long
        Dim i As Long, j As Long

        If Len(Text) =< 0 Then Exit Sub

        SpecialChars = Array("+", "*", "/", "-", "^", ":", ";", "<", ">", "=", "!", "|", "?", ",", "(", ")")
        For i = 0 To ArraySize(SpecialChars)
            Points = InStrAll(Text, CStr(SpecialChars(i)))
            Counter = 1
            If ArraySize(Points) > 0 Then
                For j = 0 To ArraySize(Points)
                    Call ColorText(Colors, ColorLengths, MidP(Text, Counter, Points(i) - 1))
                    Counter = Points(i) + 1

                    Size = ArraySize(Colors) + 1
                    ReDim Preserve Colors(Size)
                    ReDim Preserve ColorLengths(Size)
                    Colors(Size) = ColorRuleSet.GetColor(CStr(SpecialChars(i)))
                    ColorLengths(Size) = 1
                Next j
                Call ColorText(Colors, ColorLengths, MidP(Text, Counter, Len(Text)))
            End If
        Next i
    End Sub

    Private Function NoColor() As Long()
        ' For PrintConsole
    End Function
    Private Function NoColorLength() As Long()
        ' For PrintConsole
    End Function
'

' Intellisense

    Private Sub SetPositions()
        Dim Temp()       As String: Temp = Split(ConsoleText.Text, vbCrLf)
        Dim CurrentLine  As String: CurrentLine = GetLine(ConsoleText.Text, ArraySize(Temp))
        Dim WTF          As Double ' Close Enough to Multiply Lines, so that the pixeladdition for each new line is minimal. Is not a real solution to the problem but 100s or 1000s of lines can be used this way.
        WTF = 18.4375

        ScrollTop = ArraySize(Temp) * WTF - (Height / 2)
        If Len(CurrentLine) * 10 >= ConsoleText.Left + 200 Then
            ScrollLeft = Len(CurrentLine) * 10 - 200
        Else
            ScrollLeft = ConsoleText.Left
        End If
        IntellisenseList.Top = ScrollTop + (Height / 4 * 3)
        IntellisenseList.Left = ScrollLeft
        IntellisenseList.ColumnWidths = "400;1600"
    End Sub
    
    Private Sub CloseIntelliSenseList()
        IntelliSenseList.Clear
        IntelliSenseList.Visible = False
        ConsoleText.SetFocus
        ConsoleText.SelStart = Len(ConsoleText.Text)
    End Sub

    Private Sub SetUpIntelliSenseList(Text As String)
        Dim i As Long
        Dim NameSpaces As ArrayValueModel
        Set NameSpaces = Interpreter.GetIntellisenseList(Text)

        If Not NameSpaces Is Nothing Then
            For i = 0 To ArraySize(NameSpaces.Arr)
                IntelliSenseList.AddItem
                IntelliSenseList.List(IntelliSenseList.ListCount - 1, 0) = NameSpaces.Element(i).Name
                IntelliSenseList.List(IntelliSenseList.ListCount - 1, 1) = NameSpaces.Element(i).ReturnType & ", " & NameSpaces.Element(i).Arguments
            Next i
        End If
        IntelliSenseList.Visible = (IntelliSenseList.ListCount > 0)
    End Sub

    Private Sub IntelliSenseList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        
        Dim Line As String
        Dim Word As String
        Dim Words() As String
        Static IntelliSense_Index As Long

        Line = GetLine(ConsoleText.Text, CurrentLineIndex)
        If Line <> Empty Then
            Words = Split(Line, ".")
            Word  = Words(ArraySize(Words))
            Words = Split(Word, " ")
            If ArraySize(Words) <> -1 Then Word = Words(ArraySize(Words))
        End If

        Select Case KeyCode
            Case vbKeyLeft
                Call CloseIntelliSenseList()
                Exit Sub
            Case vbKeyRight
                Dim Start As Long
                Dim ReturnString As String

                If IntellisenseList.ListCount > 0 Then
                    ReturnString = IntelliSenseList.List(IntelliSense_Index, 0)
                    Start = InStr(1, ReturnString, Word)
                    If Start = 0 Then Start = 1
                    Call IConsoleView_PrintConsole(Mid(ReturnString, Start + Len(Word), Len(ReturnString)), NoColor, NoColorLength)
                    Call CloseIntelliSenseList()
                    Exit Sub
                End If
            Case vbKeyUp
                Intellisense_Index = Intellisense_Index - 1
            Case vbKeyDown
                Intellisense_Index = Intellisense_Index + 1
        End Select
        If Intellisense_Index > IntelliSenseList.ListCount - 1 Then
            Intellisense_Index = 0
        ElseIf Intellisense_Index < 0 Then
            Intellisense_Index = IntelliSenseList.ListCount - 1
        Else

        End If
        If IntelliSenseList.ListCount > 0 Then IntelliSenseList.ListIndex = IntelliSense_Index

    End Sub
'

' Error Handling
    Private Function GetError(Index As Long) As String
        Dim i As Long
        Dim Errors As ArrayValueModel
        Set Errors = Interpreter.ErrorMessages
        If Index = -1 Then
            For i = 0 To ArraySize(Errors.Arr)
                Call IConsoleView_PrintEnter(Errors.Element(i), NoColor, NoColorLength)
            Next i
        Else
            Call IConsoleView_PrintEnter(Errors.Element(Index), NoColor, NoColorLength)
        End If
    End Function
'
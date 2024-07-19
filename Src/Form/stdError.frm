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
Attribute VB_Exposed = False

Option Explicit

' Private and Public Variables
    ' Save ConsoleText.Text in Background
    Private PreviousText As String

    ' Last written line before {Enter} without Starter
    Private CurrentLine As String

    ' Index of CurrentLine
    Private CurrentLineIndex As Long

    ' Used to recognize when CurrentLine should start
    Private Const Recognizer As String = "\>>>"

    ' Determines if starter should be printed or not
    Private PasteStarter As Boolean

    ' Check if the console awaits pre-declared answer, user input or just logging
    Private WorkMode As Long
    Private Enum WorkModeEnum
        Logging = 0
        UserInputt = 1
        PreDeclaredAnswer = 2
        UserLog = 3
    End Enum

    ' Value the user put in
    Private UserInput As Variant

    ' Array of all Errormessages
        ' First Dimension is Category
        ' Second Dimension is Severity/Message
        ' Third Dimension is Index
    Private ErrorCatalog(7, 1, 99) As Variant
    ' Used once to initialize all Errormessages
    Private Initialized As Boolean

    ' Used to determine, if it is an Error
    Private Const p_IS_ERROR As Boolean = True

    ' What to show of the error
    Private p_ShowError As Boolean
    Private p_PrintError As Boolean

    ' Stops process if this Value is lower than the Error severity
    Private Const SEVERITY_BREAK As Long = 1000

    ' Runs a Question and handles Error accordingly. Should be below SEVERITY_BREAK
    Private Const ERROR_QUESTION As Long = 1

    ' Standard Value if no ErrorValue is passed
    Private Const EMPTY_ERROR As Variant = Empty

    ' De-/-activate LogMode
    Private Const LogMode As Long = 1
    Private Enum LogModeEnum
        Standard = 0    
        Console = 1
    End Enum
'







' Public Console Functions
    ' Check for UserInput
    ' Answers needs to be of same dimension as AllowedValues
    Public Function GetUserInput(Message As Variant, Optional InputType As String = "VARIANT") As Variant

        PrintConsole Message
        WorkMode = WorkModeEnum.UserInputt
        PasteStarter = False
        Do While WorkMode = WorkModeEnum.UserInputt
            DoEvents
            If UserInput <> "" Then
                UserInput = Replace(UserInput, Message, "")
                If DataType(UserInput, InputType) <> p_IS_ERROR Then
                    GetUserInput = UserInput
                    WorkMode = WorkModeEnum.Logging
                Else
                    PrintConsole Message
                End If
                UserInput = ""
            End If
        Loop
        PasteStarter = True

    End Function

    ' Check for UserInput
        ' Answers needs to be of same dimension as AllowedValues
    Public Function CheckPredeclaredAnswer(Message As Variant, AllowedValues As Variant, Optional Answers As Variant = Empty) As Variant

        Dim i As Variant
        Dim Found As Boolean
        Dim Index As Long

        Message = Message & "("
        For Each i In AllowedValues
            Message = Message & Cstr(i) & "|"
        Next i
        Message = Message & ") "
        PrintConsole Message

        WorkMode = WorkModeEnum.PreDeclaredAnswer
        PasteStarter = False
        Do While WorkMode = WorkModeEnum.PreDeclaredAnswer
            Index = 0
            DoEvents
            If UserInput <> "" Then
                UserInput = Replace(UserInput, Message, "")
                For Each i In AllowedValues
                    If Cstr(i) = UserInput Then
                        CheckPredeclaredAnswer = i
                        Found = True
                        WorkMode = WorkModeEnum.Logging
                        Exit For
                    End If
                    Index = Index + 1
                Next i
                If Found <> True Then
                    PrintEnter Handle(0, 30, UserInput)
                    PrintConsole Message
                End If
                UserInput = ""
            End If
        Loop
        PasteStarter = True
        PrintEnter Answers(Index)
        PrintConsole PrintStarter

    End Function

    Public Function PrintStarter() As Variant
        PrintStarter = ThisWorkbook.Path & Recognizer
    End Function

    Public Sub PrintEnter(Text As Variant)
        ConsoleText.Text = ConsoleText.Text & Text & Chr(13) & Chr(10)
        CurrentLineIndex = ConsoleText.LineCount
        PreviousText = ConsoleText.Text
        CurrentLine = GetLine(PreviousText, CurrentLineIndex)
    End Sub

    Public Sub PrintConsole(Text As Variant)
        ConsoleText.Text = ConsoleText.Text & Text
        CurrentLineIndex = ConsoleText.LineCount
        PreviousText = ConsoleText.Text
        CurrentLine = GetLine(PreviousText, CurrentLineIndex)
    End Sub
'

' Private Console Functions

    Private Sub UserForm_Initialize()
        ConsoleText.Text = GetStartText
        PreviousText = ConsoleText.Text
        PasteStarter = True
    End Sub

    Private Sub ConsoleText_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        Select Case KeyCode
            Case vbKeyReturn
                    HandleEnter
            Case vbKeyUp
                If CurrentLineIndex > 1 Then
                    CurrentLineIndex = CurrentLineIndex - 1
                    CurrentLine = GetLine(PreviousText, CurrentLineIndex)
                    ConsoleText.Text = PreviousText & Replace(CurrentLine, Chr(13) & Chr(10), "")
                End If
            Case vbKeyDown
                If CurrentLineIndex < ConsoleText.LineCount Then 
                    CurrentLineIndex = CurrentLineIndex + 1
                    CurrentLine = GetLine(PreviousText, CurrentLineIndex)
                    ConsoleText.Text = PreviousText & Replace(CurrentLine, Chr(13) & Chr(10), "")
                End If
        End Select
        CurrentLine = ""
    End Sub

    ' Module Code
    Private Function GetLine(Text As String, Index As Long) As String
        Dim Lines() As String
        Dim SearchString As String
        Dim ReplaceString As String
        Lines = Split(Text, vbCrLf)
        If Index > 0 And Index <= UBound(Lines) + 1 Then
            SearchString = Lines(Index - 1)
            If InStr(1, SearchString, Recognizer) = 0 Then
                ReplaceString = ""
            Else
                ReplaceString = Mid(SearchString, 1, InStr(1, SearchString, Recognizer) - 1 + Len(Recognizer))
            End If
            GetLine = Replace(SearchString, ReplaceString, "")
        Else
            GetLine = "Line number out of range"
        End If
    End Function

    Private Sub HandleEnter()
        Dim Arguments() As Variant
        Dim TempArg() As String
        Dim i As Long
        CurrentLineIndex = ConsoleText.LineCount
        PreviousText = ConsoleText.Text
        CurrentLine = GetLine(PreviousText, CurrentLineIndex - 1)

        Select Case WorkMode
            Case WorkModeEnum.Logging
                If InStr(1, CurrentLine, "; ") <> 0 Then
                    TempArg = Split(CurrentLine, "; ")
                    ReDim Arguments(Ubound(TempArg)) As Variant
                Else
                    ReDim TempArg(0) As String
                    ReDim Arguments(0) As Variant
                    TempArg(0) = CurrentLine
                End If
                For i = 0 To Ubound(TempArg)
                    Arguments(i) = CStr(TempArg(i))
                Next i
                PrintEnter HandleCode(Arguments)
            Case WorkModeEnum.UserInputt, WorkModeEnum.PredeclaredAnswer
                UserInput = Replace(CurrentLine, Chr(13) & Chr(10), "")
            Case WorkModeEnum.UserLog
        End Select
        If PasteStarter = True Then PrintConsole PrintStarter
    End Sub

    Private Function HandleCode(ParamArray Arguments() As Variant) As Variant

        On Error GoTo Error
        If HandleSpecial(Arguments) = Empty Then HandleCode = "Success": Exit Function
        Select Case UBound(Arguments(0))
            Case 00:   HandleCode = Application.Run(Arguments(0)(0))
            Case 01:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1))
            Case 02:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2))
            Case 03:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3))
            Case 04:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4))
            Case 05:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5))
            Case 06:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6))
            Case 07:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7))
            Case 08:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8))
            Case 09:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9))
            Case 10:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10))
            Case 11:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11))
            Case 12:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12))
            Case 13:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13))
            Case 14:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14))
            Case 15:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15))
            Case 16:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16))
            Case 17:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17))
            Case 18:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18))
            Case 19:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19))
            Case 20:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20))
            Case 21:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21))
            Case 22:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22))
            Case 23:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22), Arguments(0)(23))
            Case 24:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22), Arguments(0)(23), Arguments(0)(24))
            Case 25:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22), Arguments(0)(23), Arguments(0)(24), Arguments(0)(25))
            Case 26:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22), Arguments(0)(23), Arguments(0)(24), Arguments(0)(25), Arguments(0)(26))
            Case 27:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22), Arguments(0)(23), Arguments(0)(24), Arguments(0)(25), Arguments(0)(26), Arguments(0)(27))
            Case 28:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22), Arguments(0)(23), Arguments(0)(24), Arguments(0)(25), Arguments(0)(26), Arguments(0)(27), Arguments(0)(28))
            Case 29:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22), Arguments(0)(23), Arguments(0)(24), Arguments(0)(25), Arguments(0)(26), Arguments(0)(27), Arguments(0)(28), Arguments(0)(29))
            Case 30:   HandleCode = Application.Run(Arguments(0)(0), Arguments(0)(1), Arguments(0)(2), Arguments(0)(3), Arguments(0)(4), Arguments(0)(5), Arguments(0)(6), Arguments(0)(7), Arguments(0)(8), Arguments(0)(9), Arguments(0)(10), Arguments(0)(11), Arguments(0)(12), Arguments(0)(13), Arguments(0)(14), Arguments(0)(15), Arguments(0)(16), Arguments(0)(17), Arguments(0)(18), Arguments(0)(19), Arguments(0)(20), Arguments(0)(21), Arguments(0)(22), Arguments(0)(23), Arguments(0)(24), Arguments(0)(25), Arguments(0)(26), Arguments(0)(27), Arguments(0)(28), Arguments(0)(29), Arguments(0)(30))
            Case Else: HandleCode = "Too many Arguments"
        End Select
            If HandleCode = Empty Then HandleCode = "Success"
            Exit Function
        Error:
        HandleCode = "Could not run Procedure. Procedure might not exist"

    End Function

    Private Function HandleSpecial(ParamArray Arguments() As Variant) As Variant

        Select Case UCase(CStr(Arguments(0)(0)(0)))
        Case "HELP": HandleHelp
        Case "CLEAR":
        Case "":
        Case Else: HandleSpecial = 1
        End Select

    End Function

    Private Function GetStartText() As String
        GetStartText =                   _
        "VBA Console [Version 1.0]" & Chr(13) & Chr(10) & _
        "No Rights reserved"        & Chr(13) & Chr(10) & _
        Chr(13) & Chr(10)                               & _
        PrintStarter
    End Function

    Private Sub HandleHelp()

        Dim Message As String
        Message = _
        "--------------------------------------------------"                                           & Chr(13) & Chr(10) & _
        "This Console can do the following:"                                                           & Chr(13) & Chr(10) & _
        "1. It can be used as a form to showw messages, ask questions to the user or get a user input" & Chr(13) & Chr(10) & _
        "2. It can be used to showw and log errors and handle them by user input"                      & Chr(13) & Chr(10) & _
        "3. It can run Procedures with up to 29 arguments"                                             & Chr(13) & Chr(10) & _
        ""                                                                                             & Chr(13) & Chr(10) & _
        "HOW TO USE IT:"                                                                               & Chr(13) & Chr(10) & _
        "   Run a Procedure:"                                                                          & Chr(13) & Chr(10) & _
        "       To run a procedure you have to write the name of said procedure (Case sensitive)"      & Chr(13) & Chr(10) & _
        "       If you want to pass parameters you have to write    |; | between every parameter"      & Chr(13) & Chr(10) & _
        "       Example:"                                                                              & Chr(13) & Chr(10) & _
        "           Say; THIS IS A PARAMETER; THIS IS ANOTHER PARAMETER"                               & Chr(13) & Chr(10) & _
        ""                                                                                             & Chr(13) & Chr(10) & _
        "   Ask a question:"                                                                           & Chr(13) & Chr(10) & _
        "       Use CheckPredeclaredAnswer"                                                            & Chr(13) & Chr(10) & _
        "           Param1 = Message to be showwn"                                                     & Chr(13) & Chr(10) & _
        "           Param2 = Array of Values, which are acceptable answers"                            & Chr(13) & Chr(10) & _
        "           Param3 = Array of Messages, which showw a text according to answer in Param2"      & Chr(13) & Chr(10) & _
        "       The Function will loop until one of the acceptable answers is typed"                   & Chr(13) & Chr(10) & _
        "--------------------------------------------------"                                           & Chr(13) & Chr(10)
        PrintEnter Message

    End Sub
'






' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------







' Public Errormethods
    Public Property Get IS_ERROR()
        IS_ERROR = p_IS_ERROR
    End Property
        
    Public Property Let ShowError(Value As Boolean)
        p_ShowError = Value
    End Property
    Public Property Let PrintError(Value As Boolean)
        p_PrintError = Value
    End Property


    ' Handles Errormessage
    Public Function Handle(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR) As Boolean

        Dim Severity As Integer

        ProtInit
        Severity = ErrorCatalog(ErrorCategory, 0, ErrorIndex)
        If Severity = ERROR_QUESTION Then
            Handle = Ask(ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4)
        Else
            Showw ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4
            Handle = IS_ERROR
        End If
        Printt ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4 
        If Severity > SEVERITY_BREAK Then
            End
        End If

    End Function

    ' Compare Numbers or Text and Handle Errors
    Public Function Variable(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean

        ProtInit
        If Operator <> Empty Then
            Select Case UCase(Operator)
                Case "=", "IS"
                    If FirstValue =  SecondValue Then Variable = Handle(0, 15, FirstValue, SecondValue): Exit Function
                Case "<>", "NOT", "!="
                    If FirstValue <> SecondValue Then Variable = Handle(0, 06, FirstValue, SecondValue): Exit Function
                Case "<"
                    If FirstValue <  SecondValue Then Variable = Handle(0, 09, FirstValue, SecondValue): Exit Function
                Case ">"
                    If FirstValue >  SecondValue Then Variable = Handle(0, 10, FirstValue, SecondValue): Exit Function
                Case "=<", "<="
                    If FirstValue =< SecondValue Then Variable = Handle(0, 07, FirstValue, SecondValue): Exit Function
                Case ">=", "=>"
                    If FirstValue >= SecondValue Then Variable = Handle(0, 08, FirstValue, SecondValue): Exit Function
                Case Else
            End Select
        End If
        If MinValue <> Empty Then
            If FirstValue < MinValue             Then Variable = Handle(0, 09, FirstValue, MinValue):    Exit Function
        End If
        If MaxValue <> Empty Then
            If FirstValue > MaxValue             Then Variable = Handle(0, 10, FirstValue, MaxValue):    Exit Function
        End If
        If FirstValue = Empty                    Then Variable = Handle(0, 02, "FirstValue"):            Exit Function

    End Function

    ' Compare Numbers or Text and Handle Errors
    Public Function Object(FirstValue As Object, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Object = EMPTY_ERROR) As Boolean

        ProtInit
        If Operator <> Empty Then
            Select Case UCase(Operator)
                Case "=", "IS"
                    If FirstValue     IS SecondValue Then Object = Handle(0, 15, "FirstValue", "SecondValue"): Exit Function
                Case "<>", "NOT", "!="
                    If Not FirstValue IS SecondValue Then Object = Handle(0, 06, "FirstValue", "SecondValue"): Exit Function
            End Select
        End If
        If FirstValue Is Nothing                      Then Object = Handle(0, 03, "FirstValue"):               Exit Function

    End Function

    ' Handles possible Errors when working with workbooks
    Public Function Workbook(WorkbookName As String, Optional ShouldExist As Boolean = True) As Boolean

        Dim WB As WorkBook
        ProtInit
        If WorkbookName = Empty Then Workbook = Handle(0, 02, "WorkbookName"): Exit Function
        If ShouldExist = True Then
            For Each WB in WorkBooks
                If WB.Name = WorkbookName Then Exit Function
            Next
            Workbook = Handle(3, 00, WorkbookName)
        Else
            For Each WB in WorkBooks
                If WB.Name = WorkbookName Then Workbook = Handle(3, 01, WorkbookName): Exit Function
            Next
        End If

    End Function

    ' Handles possible Errors when working with workheets
    Public Function Worksheet(WorkbookName As String, SheetName As String, Optional ShouldExist As Boolean = True) As Boolean

        Dim WS As Worksheet
        ProtInit
        If Workbook(WorkBookName, True) = IS_ERROR Then Worksheet = IS_ERROR: Exit Function
        If SheetName = Empty Then Worksheet = Handle(0, 02, "SheetName"   ):  Exit Function
        With Workbooks(WorkbookName)
            If ShouldExist = True Then
                For Each WS in .Worksheets
                    If WS.Name = SheetName Then Exit Function
                Next
                Worksheet = Handle(4, 01, WorkbookName, SheetName)
            Else
                For Each WS in .Worksheets
                    If WS.Name = SheetName Then Worksheet = Handle(4, 00, WorkbookName, SheetName): Exit Function
                Next
            End If
        End With

    End Function

    ' Compare Strings and Handle Errors
    Public Function Strings(Text As String, Operator As String, SecondText As String) As Boolean
    
        ProtInit
        If Operator <> Empty Then
            Select Case UCase(Operator)
                Case "=", "IS"
                    If Text =        SecondText Then Strings = Handle(0, 15, Text, SecondText): Exit Function
                Case "<>", "NOT", "!="
                    If Text <>       SecondText Then Strings = Handle(0, 06, Text, SecondText): Exit Function
                Case "<"
                    If Text <        SecondText Then Strings = Handle(0, 10, Text, SecondText): Exit Function
                Case ">"
                    If Text >        SecondText Then Strings = Handle(0, 09, Text, SecondText): Exit Function
                Case "=<", "<="
                    If Text =<       SecondText Then Strings = Handle(0, 08, Text, SecondText): Exit Function
                Case ">=", "=>"
                    If Text >=       SecondText Then Strings = Handle(0, 07, Text, SecondText): Exit Function
                Case "LIKE"
                    If Text Like     SecondText Then Strings = Handle(0, 17, Text, SecondText): Exit Function
                Case "NOT LIKE", "UNLIKE"
                    If Not Text Like SecondText Then Strings = Handle(0, 16, Text, SecondText): Exit Function
                Case Else
            End Select
        End If
        If Text = Empty                         Then Strings = Handle(0, 02, "Text"):           Exit Function

    End Function

    ' Superset of Variable, used to check if it is a number and Handle Errors
    Public Function Number(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean
        ProtInit
        If IsNumeric(FirstValue) = False Then
            Number = Handle(0, 18, FirstValue)
            Exit Function
        Else
            Number = Variable(FirstValue, Operator, SecondValue, MinValue, MaxValue)
        End If
    End Function

    ' Superset of Variable, used to check if it is a date and Handle Errors
    Public Function Dates(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean
        ProtInit
        If IsDate(FirstValue) = False Then
            Dates = Handle(0, 18, Text, SecondText)
            Exit Function
        Else
            Dates = Variable(FirstValue, Operator, SecondValue, MinValue, MaxValue)
        End If
    End Function

    ' Handles File Validation and Handle Errors
    Public Function File(FilePath As String, Optional ShouldExist As Boolean = True) As Boolean

        If ShouldExist = True Then
            If Dir(FilePath) = "" Then
                File = Handle(0, 19, FilePath)
            End If
        Else
            If Dir(FilePath) <> "" Then
                File = Handle(0, 20, FilePath)
            End If
        End If   
        
    End Function

    ' Check Connection to specified Computer and Handle Errors
    Public Function Connection(Optional Computer As String = ".", Optional ShouldExist As Boolean = True) As Boolean

        Dim objWMIService As Object
        Dim colItems As Object
        Dim objItem As Object
        
        Set objWMIService = GetObject("winmgmts:\\" & Computer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
        If ShouldExist = True Then
            If colItems.Count = 0 Then
                Connection = Handle(0, 21, Computer)
            Else
                For Each objItem In colItems
                    If objItem.IPAddress(0) <> "" Then
                        Exit Function
                    End If
                Next
                Connection = Handle(0, 22, Computer)
            End If
        Else
            If colItems.Count <> 0 Then
                Connection = Handle(0, 24, Computer)
            Else
                For Each objItem In colItems
                    If objItem.IPAddress(0) = "" Then
                        Exit Function
                    End If
                Next
                Connection = Handle(0, 25, Computer)
            End If
        End If

    End Function

    ' Check DatabaseConnection and Handle Errors
    Public Function ConnectToDatabase(DataBasePath As String) As Boolean

        On Error GoTo ErrorHandler
        Dim Conn As Object
        Set Conn = CreateObject("ADODB.Connection")
        Conn.Open DataBasePath
        If Conn.State = 1 Then
            ConnectToDatabase = True
        Else
            ConnectToDatabase = Handle(5, 26, DataBasePath)
        End If
        Exit Function
        ErrorHandler:
        ConnectToDatabase = Handle(5, 23, "ConnectToDatabase", DataBasePath)

    End Function

    Public Function DataType(Value As Variant, InputType As String, Optional ShouldBe As Boolean = True) As Boolean
        
        Dim Inputt As Boolean
        
        Select Case UCase(InputType)
            Case "VARIANT":         If VarType(Value)                                                 = vbVariant  Then Inputt = True
            Case "STRING":          If VarType(Value)                                                 = vbString   Then Inputt = True
            Case "NUMBER":          If IsNumeric(Value)                                               = True       Then Inputt = True
            Case "DATE":            If IsDate(Value)                                                  = True       Then Inputt = True
            Case "BOOLEAN":         If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbBoolean  Then Inputt = True
            Case "INTEGER", "INT":  If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbInteger  Then Inputt = True
            Case "LONG":            If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbLong     Then Inputt = True
            Case "LONGLONG":        If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbLongLong Then Inputt = True
            Case "DOUBLE":          If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbDouble   Then Inputt = True
            Case "SINGLE", "FLOAT": If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbSingle   Then Inputt = True
            Case "EMPTY":           If IsEmpty(Value)                                                 = True       Then Inputt = True
            Case "NULL":            If IsNull(Value)                                                  = True       Then Inputt = True
            Case Else:             DataType = Handle(0, 29, Value, InputType)
        End Select
        If Not Inputt Xor True Then
            Exit Function
        Else
            If Inputt = True Then
                DataType = Handle(0, 28, Value, InputType)
            Else
                DataType = Handle(0, 27, Value, InputType)
            End If
        End If        
        Exit Function

    End Function
'



' Private Errormethods
    ' Print Error to Immediate
    Private Sub Printt(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR)
        
        Select Case LogMode
            Case Standard: Debug.Print GetMessage(ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4)
            Case Console: PrintEnter GetMessage(ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4)
        End Select

    End Sub
    
    ' Print Error as MessageBox
    Private Sub Showw(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR)
        
        Dim Temp As Variant
        Select Case LogMode
            Case Standard: Temp = MsgBox(GetMessage(ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4), vbExclamation, "ERROR")
            Case Console: ' Nothing, as it happens in Printt for the console
        End Select
        
    End Sub
    
    ' Asks Yes/No Question (No will raise an Error)
    Private Function Ask(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR) As Boolean
        Dim Temp As Variant
        Dim ArrayAnswer() As Variant
        Dim ArrayMessage() As Variant
        ArrayAnswer = Array("y","n")
        ArrayMessage = Array("Answer is Yes", "Answer is No")
        Select Case LogMode
            Case Standard: Temp = MsgBox(GetMessage(ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4), vbYesNo, "QUESTION")
            Case Console: Temp = CheckInput(GetMessage(ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4), ArrayAnswer, ArrayMessage)
        End Select
        If Temp = False Then Ask = IS_ERROR
    End Function
    
    ' Gets Errormessage
    Private Function GetMessage(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR) As String

        Dim ErrorMessage As String
        Dim String1 As String
        String1 = ErrorCatalog(ErrorCategory, 1, ErrorIndex)
        ProtInit
        ErrorMessage = "Category: " & GetCategory(ErrorCategory)                                     & Chr(13) & _
                       "Severity: " & ErrorCatalog(ErrorCategory, 0, ErrorIndex)                     & Chr(13) & _
                       "Index   : " & ErrorIndex                                                     & Chr(13) & _
                       "Message : " & String1                                                        & Chr(13)
        If ErrorValue1 <> EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "Value1  : " & ErrorValue1 & Chr(13)
        If ErrorValue2 <> EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "Value2  : " & ErrorValue2 & Chr(13)
        If ErrorValue3 <> EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "Value3  : " & ErrorValue3 & Chr(13)
        If ErrorValue4 <> EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "Value4  : " & ErrorValue4 & Chr(13)
        ErrorMessage = ErrorMessage & "------------------------------------------------------------------------------"
        GetMessage = ErrorMessage

    End Function

    ' Gets Name of Errorcategory
    Private Function GetCategory(ErrorCategory As Long) As String

        Select Case ErrorCategory
            Case 0: GetCategory = "System"
            Case 1: GetCategory = "Workbook"
            Case 2: GetCategory = "Worksheet"
            Case 3: GetCategory = "Linker"
            Case 4: GetCategory = "Compiler"
            Case 5: GetCategory = "Module"
            Case 6: GetCategory = "Class"
            Case 7: GetCategory = "Userform"
            Case Else
                    GetCategory = "UNKNOWN"
        End Select

    End Function

    ' Runs once to Initialize all Errormessages
    Private Sub ProtInit()

        If Initialized = False Then
        ' System-Errors
            ErrorCatalog(0000, 0000, 0000) = 1000: ErrorCatalog(0000, 0001, 0000) = "ErrorCategory doesnt exist"
            ErrorCatalog(0000, 0000, 0001) = 1000: ErrorCatalog(0000, 0001, 0001) = "Value isnt available"
            ErrorCatalog(0000, 0000, 0002) = 1000: ErrorCatalog(0000, 0001, 0002) = "Value is Empty"
            ErrorCatalog(0000, 0000, 0003) = 1000: ErrorCatalog(0000, 0001, 0003) = "Value is Nothing"
            ErrorCatalog(0000, 0000, 0004) = 1000: ErrorCatalog(0000, 0001, 0004) = "Value Overflow"
            ErrorCatalog(0000, 0000, 0005) = 1000: ErrorCatalog(0000, 0001, 0005) = "Value Underflow"
            ErrorCatalog(0000, 0000, 0006) = 1000: ErrorCatalog(0000, 0001, 0006) = "Value1 doesnt equal Value2"
            ErrorCatalog(0000, 0000, 0007) = 1000: ErrorCatalog(0000, 0001, 0007) = "Value1 is smaller than or equal to Value2"
            ErrorCatalog(0000, 0000, 0008) = 1000: ErrorCatalog(0000, 0001, 0008) = "Value1 is bigger than or equal to Value2"
            ErrorCatalog(0000, 0000, 0009) = 1000: ErrorCatalog(0000, 0001, 0009) = "Value1 is smaller than Value2"
            ErrorCatalog(0000, 0000, 0010) = 1000: ErrorCatalog(0000, 0001, 0010) = "Value1 is bigger than Value2"
            ErrorCatalog(0000, 0000, 0011) = 1000: ErrorCatalog(0000, 0001, 0011) = "Value1 is Value2"
            ErrorCatalog(0000, 0000, 0012) = 1000: ErrorCatalog(0000, 0001, 0012) = "Several Values are Empty"
            ErrorCatalog(0000, 0000, 0013) = 1000: ErrorCatalog(0000, 0001, 0013) = "To many Values arent Empty"
            ErrorCatalog(0000, 0000, 0014) = 1000: ErrorCatalog(0000, 0001, 0014) = "Value Is Something"
            ErrorCatalog(0000, 0000, 0015) = 1000: ErrorCatalog(0000, 0001, 0015) = "Value1 equals Value2"
            ErrorCatalog(0000, 0000, 0016) = 1000: ErrorCatalog(0000, 0001, 0016) = "Value1 is not like Value2"
            ErrorCatalog(0000, 0000, 0017) = 1000: ErrorCatalog(0000, 0001, 0017) = "Value1 is like Value2"
            ErrorCatalog(0000, 0000, 0018) = 1000: ErrorCatalog(0000, 0001, 0018) = "Value is not a Number"
            ErrorCatalog(0000, 0000, 0019) = 1000: ErrorCatalog(0000, 0001, 0019) = "File does not exist"
            ErrorCatalog(0000, 0000, 0020) = 1000: ErrorCatalog(0000, 0001, 0020) = "File exists"
            ErrorCatalog(0000, 0000, 0021) = 1000: ErrorCatalog(0000, 0001, 0021) = "No active network Connection"
            ErrorCatalog(0000, 0000, 0022) = 1000: ErrorCatalog(0000, 0001, 0022) = "No valid IP address found"
            ErrorCatalog(0000, 0000, 0023) = 1000: ErrorCatalog(0000, 0001, 0023) = "Unknown error"
            ErrorCatalog(0000, 0000, 0024) = 1000: ErrorCatalog(0000, 0001, 0024) = "Active network Connection"
            ErrorCatalog(0000, 0000, 0025) = 1000: ErrorCatalog(0000, 0001, 0025) = "valid IP address found"
            ErrorCatalog(0000, 0000, 0026) = 1000: ErrorCatalog(0000, 0001, 0026) = "Unable to open Connection"
            ErrorCatalog(0000, 0000, 0027) = 1000: ErrorCatalog(0000, 0001, 0027) = "Should not be this Datatype"
            ErrorCatalog(0000, 0000, 0028) = 1000: ErrorCatalog(0000, 0001, 0028) = "Should be this Datatype"
            ErrorCatalog(0000, 0000, 0029) = 1000: ErrorCatalog(0000, 0001, 0029) = "Unknown Datatype"
            ErrorCatalog(0000, 0000, 0030) = 1000: ErrorCatalog(0000, 0001, 0030) = "Value is not Valid"
            ErrorCatalog(0000, 0000, 0031) = 1000: ErrorCatalog(0000, 0001, 0031) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0032) = 1000: ErrorCatalog(0000, 0001, 0032) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0033) = 1000: ErrorCatalog(0000, 0001, 0033) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0034) = 1000: ErrorCatalog(0000, 0001, 0034) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0035) = 1000: ErrorCatalog(0000, 0001, 0035) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0036) = 1000: ErrorCatalog(0000, 0001, 0036) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0037) = 1000: ErrorCatalog(0000, 0001, 0037) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0038) = 1000: ErrorCatalog(0000, 0001, 0038) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0039) = 1000: ErrorCatalog(0000, 0001, 0039) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0040) = 1000: ErrorCatalog(0000, 0001, 0040) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0041) = 1000: ErrorCatalog(0000, 0001, 0041) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0042) = 1000: ErrorCatalog(0000, 0001, 0042) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0043) = 1000: ErrorCatalog(0000, 0001, 0043) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0044) = 1000: ErrorCatalog(0000, 0001, 0044) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0045) = 1000: ErrorCatalog(0000, 0001, 0045) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0046) = 1000: ErrorCatalog(0000, 0001, 0046) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0047) = 1000: ErrorCatalog(0000, 0001, 0047) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0048) = 1000: ErrorCatalog(0000, 0001, 0048) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0049) = 1000: ErrorCatalog(0000, 0001, 0049) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0050) = 1000: ErrorCatalog(0000, 0001, 0050) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0051) = 1000: ErrorCatalog(0000, 0001, 0051) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0052) = 1000: ErrorCatalog(0000, 0001, 0052) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0053) = 1000: ErrorCatalog(0000, 0001, 0053) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0054) = 1000: ErrorCatalog(0000, 0001, 0054) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0055) = 1000: ErrorCatalog(0000, 0001, 0055) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0056) = 1000: ErrorCatalog(0000, 0001, 0056) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0057) = 1000: ErrorCatalog(0000, 0001, 0057) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0058) = 1000: ErrorCatalog(0000, 0001, 0058) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0059) = 1000: ErrorCatalog(0000, 0001, 0059) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0060) = 1000: ErrorCatalog(0000, 0001, 0060) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0061) = 1000: ErrorCatalog(0000, 0001, 0061) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0062) = 1000: ErrorCatalog(0000, 0001, 0062) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0063) = 1000: ErrorCatalog(0000, 0001, 0063) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0064) = 1000: ErrorCatalog(0000, 0001, 0064) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0065) = 1000: ErrorCatalog(0000, 0001, 0065) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0066) = 1000: ErrorCatalog(0000, 0001, 0066) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0067) = 1000: ErrorCatalog(0000, 0001, 0067) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0068) = 1000: ErrorCatalog(0000, 0001, 0068) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0069) = 1000: ErrorCatalog(0000, 0001, 0069) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0070) = 1000: ErrorCatalog(0000, 0001, 0070) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0071) = 1000: ErrorCatalog(0000, 0001, 0071) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0072) = 1000: ErrorCatalog(0000, 0001, 0072) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0073) = 1000: ErrorCatalog(0000, 0001, 0073) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0074) = 1000: ErrorCatalog(0000, 0001, 0074) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0075) = 1000: ErrorCatalog(0000, 0001, 0075) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0076) = 1000: ErrorCatalog(0000, 0001, 0076) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0077) = 1000: ErrorCatalog(0000, 0001, 0077) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0078) = 1000: ErrorCatalog(0000, 0001, 0078) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0079) = 1000: ErrorCatalog(0000, 0001, 0079) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0080) = 1000: ErrorCatalog(0000, 0001, 0080) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0081) = 1000: ErrorCatalog(0000, 0001, 0081) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0082) = 1000: ErrorCatalog(0000, 0001, 0082) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0083) = 1000: ErrorCatalog(0000, 0001, 0083) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0084) = 1000: ErrorCatalog(0000, 0001, 0084) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0085) = 1000: ErrorCatalog(0000, 0001, 0085) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0086) = 1000: ErrorCatalog(0000, 0001, 0086) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0087) = 1000: ErrorCatalog(0000, 0001, 0087) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0088) = 1000: ErrorCatalog(0000, 0001, 0088) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0089) = 1000: ErrorCatalog(0000, 0001, 0089) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0090) = 1000: ErrorCatalog(0000, 0001, 0090) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0091) = 1000: ErrorCatalog(0000, 0001, 0091) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0092) = 1000: ErrorCatalog(0000, 0001, 0092) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0093) = 1000: ErrorCatalog(0000, 0001, 0093) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0094) = 1000: ErrorCatalog(0000, 0001, 0094) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0095) = 1000: ErrorCatalog(0000, 0001, 0095) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0096) = 1000: ErrorCatalog(0000, 0001, 0096) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0097) = 1000: ErrorCatalog(0000, 0001, 0097) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0098) = 1000: ErrorCatalog(0000, 0001, 0098) = "PLACEHOLDER"
            ErrorCatalog(0000, 0000, 0099) = 1000: ErrorCatalog(0000, 0001, 0099) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0000) = 1000: ErrorCatalog(0001, 0001, 0000) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0001) = 1000: ErrorCatalog(0001, 0001, 0001) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0002) = 1000: ErrorCatalog(0001, 0001, 0002) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0003) = 1000: ErrorCatalog(0001, 0001, 0003) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0004) = 1000: ErrorCatalog(0001, 0001, 0004) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0005) = 1000: ErrorCatalog(0001, 0001, 0005) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0006) = 1000: ErrorCatalog(0001, 0001, 0006) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0007) = 1000: ErrorCatalog(0001, 0001, 0007) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0008) = 1000: ErrorCatalog(0001, 0001, 0008) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0009) = 1000: ErrorCatalog(0001, 0001, 0009) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0010) = 1000: ErrorCatalog(0001, 0001, 0010) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0011) = 1000: ErrorCatalog(0001, 0001, 0011) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0012) = 1000: ErrorCatalog(0001, 0001, 0012) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0013) = 1000: ErrorCatalog(0001, 0001, 0013) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0014) = 1000: ErrorCatalog(0001, 0001, 0014) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0015) = 1000: ErrorCatalog(0001, 0001, 0015) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0016) = 1000: ErrorCatalog(0001, 0001, 0016) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0017) = 1000: ErrorCatalog(0001, 0001, 0017) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0018) = 1000: ErrorCatalog(0001, 0001, 0018) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0019) = 1000: ErrorCatalog(0001, 0001, 0019) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0020) = 1000: ErrorCatalog(0001, 0001, 0020) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0021) = 1000: ErrorCatalog(0001, 0001, 0021) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0022) = 1000: ErrorCatalog(0001, 0001, 0022) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0023) = 1000: ErrorCatalog(0001, 0001, 0023) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0024) = 1000: ErrorCatalog(0001, 0001, 0024) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0025) = 1000: ErrorCatalog(0001, 0001, 0025) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0026) = 1000: ErrorCatalog(0001, 0001, 0026) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0027) = 1000: ErrorCatalog(0001, 0001, 0027) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0028) = 1000: ErrorCatalog(0001, 0001, 0028) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0029) = 1000: ErrorCatalog(0001, 0001, 0029) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0030) = 1000: ErrorCatalog(0001, 0001, 0030) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0031) = 1000: ErrorCatalog(0001, 0001, 0031) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0032) = 1000: ErrorCatalog(0001, 0001, 0032) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0033) = 1000: ErrorCatalog(0001, 0001, 0033) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0034) = 1000: ErrorCatalog(0001, 0001, 0034) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0035) = 1000: ErrorCatalog(0001, 0001, 0035) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0036) = 1000: ErrorCatalog(0001, 0001, 0036) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0037) = 1000: ErrorCatalog(0001, 0001, 0037) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0038) = 1000: ErrorCatalog(0001, 0001, 0038) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0039) = 1000: ErrorCatalog(0001, 0001, 0039) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0040) = 1000: ErrorCatalog(0001, 0001, 0040) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0041) = 1000: ErrorCatalog(0001, 0001, 0041) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0042) = 1000: ErrorCatalog(0001, 0001, 0042) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0043) = 1000: ErrorCatalog(0001, 0001, 0043) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0044) = 1000: ErrorCatalog(0001, 0001, 0044) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0045) = 1000: ErrorCatalog(0001, 0001, 0045) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0046) = 1000: ErrorCatalog(0001, 0001, 0046) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0047) = 1000: ErrorCatalog(0001, 0001, 0047) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0048) = 1000: ErrorCatalog(0001, 0001, 0048) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0049) = 1000: ErrorCatalog(0001, 0001, 0049) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0050) = 1000: ErrorCatalog(0001, 0001, 0050) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0051) = 1000: ErrorCatalog(0001, 0001, 0051) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0052) = 1000: ErrorCatalog(0001, 0001, 0052) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0053) = 1000: ErrorCatalog(0001, 0001, 0053) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0054) = 1000: ErrorCatalog(0001, 0001, 0054) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0055) = 1000: ErrorCatalog(0001, 0001, 0055) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0056) = 1000: ErrorCatalog(0001, 0001, 0056) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0057) = 1000: ErrorCatalog(0001, 0001, 0057) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0058) = 1000: ErrorCatalog(0001, 0001, 0058) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0059) = 1000: ErrorCatalog(0001, 0001, 0059) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0060) = 1000: ErrorCatalog(0001, 0001, 0060) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0061) = 1000: ErrorCatalog(0001, 0001, 0061) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0062) = 1000: ErrorCatalog(0001, 0001, 0062) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0063) = 1000: ErrorCatalog(0001, 0001, 0063) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0064) = 1000: ErrorCatalog(0001, 0001, 0064) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0065) = 1000: ErrorCatalog(0001, 0001, 0065) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0066) = 1000: ErrorCatalog(0001, 0001, 0066) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0067) = 1000: ErrorCatalog(0001, 0001, 0067) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0068) = 1000: ErrorCatalog(0001, 0001, 0068) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0069) = 1000: ErrorCatalog(0001, 0001, 0069) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0070) = 1000: ErrorCatalog(0001, 0001, 0070) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0071) = 1000: ErrorCatalog(0001, 0001, 0071) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0072) = 1000: ErrorCatalog(0001, 0001, 0072) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0073) = 1000: ErrorCatalog(0001, 0001, 0073) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0074) = 1000: ErrorCatalog(0001, 0001, 0074) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0075) = 1000: ErrorCatalog(0001, 0001, 0075) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0076) = 1000: ErrorCatalog(0001, 0001, 0076) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0077) = 1000: ErrorCatalog(0001, 0001, 0077) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0078) = 1000: ErrorCatalog(0001, 0001, 0078) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0079) = 1000: ErrorCatalog(0001, 0001, 0079) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0080) = 1000: ErrorCatalog(0001, 0001, 0080) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0081) = 1000: ErrorCatalog(0001, 0001, 0081) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0082) = 1000: ErrorCatalog(0001, 0001, 0082) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0083) = 1000: ErrorCatalog(0001, 0001, 0083) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0084) = 1000: ErrorCatalog(0001, 0001, 0084) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0085) = 1000: ErrorCatalog(0001, 0001, 0085) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0086) = 1000: ErrorCatalog(0001, 0001, 0086) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0087) = 1000: ErrorCatalog(0001, 0001, 0087) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0088) = 1000: ErrorCatalog(0001, 0001, 0088) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0089) = 1000: ErrorCatalog(0001, 0001, 0089) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0090) = 1000: ErrorCatalog(0001, 0001, 0090) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0091) = 1000: ErrorCatalog(0001, 0001, 0091) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0092) = 1000: ErrorCatalog(0001, 0001, 0092) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0093) = 1000: ErrorCatalog(0001, 0001, 0093) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0094) = 1000: ErrorCatalog(0001, 0001, 0094) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0095) = 1000: ErrorCatalog(0001, 0001, 0095) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0096) = 1000: ErrorCatalog(0001, 0001, 0096) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0097) = 1000: ErrorCatalog(0001, 0001, 0097) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0098) = 1000: ErrorCatalog(0001, 0001, 0098) = "PLACEHOLDER"
            ErrorCatalog(0001, 0000, 0099) = 1000: ErrorCatalog(0001, 0001, 0099) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0000) = 1000: ErrorCatalog(0002, 0001, 0000) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0001) = 1000: ErrorCatalog(0002, 0001, 0001) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0002) = 1000: ErrorCatalog(0002, 0001, 0002) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0003) = 1000: ErrorCatalog(0002, 0001, 0003) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0004) = 1000: ErrorCatalog(0002, 0001, 0004) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0005) = 1000: ErrorCatalog(0002, 0001, 0005) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0006) = 1000: ErrorCatalog(0002, 0001, 0006) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0007) = 1000: ErrorCatalog(0002, 0001, 0007) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0008) = 1000: ErrorCatalog(0002, 0001, 0008) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0009) = 1000: ErrorCatalog(0002, 0001, 0009) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0010) = 1000: ErrorCatalog(0002, 0001, 0010) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0011) = 1000: ErrorCatalog(0002, 0001, 0011) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0012) = 1000: ErrorCatalog(0002, 0001, 0012) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0013) = 1000: ErrorCatalog(0002, 0001, 0013) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0014) = 1000: ErrorCatalog(0002, 0001, 0014) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0015) = 1000: ErrorCatalog(0002, 0001, 0015) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0016) = 1000: ErrorCatalog(0002, 0001, 0016) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0017) = 1000: ErrorCatalog(0002, 0001, 0017) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0018) = 1000: ErrorCatalog(0002, 0001, 0018) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0019) = 1000: ErrorCatalog(0002, 0001, 0019) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0020) = 1000: ErrorCatalog(0002, 0001, 0020) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0021) = 1000: ErrorCatalog(0002, 0001, 0021) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0022) = 1000: ErrorCatalog(0002, 0001, 0022) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0023) = 1000: ErrorCatalog(0002, 0001, 0023) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0024) = 1000: ErrorCatalog(0002, 0001, 0024) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0025) = 1000: ErrorCatalog(0002, 0001, 0025) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0026) = 1000: ErrorCatalog(0002, 0001, 0026) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0027) = 1000: ErrorCatalog(0002, 0001, 0027) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0028) = 1000: ErrorCatalog(0002, 0001, 0028) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0029) = 1000: ErrorCatalog(0002, 0001, 0029) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0030) = 1000: ErrorCatalog(0002, 0001, 0030) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0031) = 1000: ErrorCatalog(0002, 0001, 0031) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0032) = 1000: ErrorCatalog(0002, 0001, 0032) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0033) = 1000: ErrorCatalog(0002, 0001, 0033) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0034) = 1000: ErrorCatalog(0002, 0001, 0034) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0035) = 1000: ErrorCatalog(0002, 0001, 0035) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0036) = 1000: ErrorCatalog(0002, 0001, 0036) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0037) = 1000: ErrorCatalog(0002, 0001, 0037) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0038) = 1000: ErrorCatalog(0002, 0001, 0038) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0039) = 1000: ErrorCatalog(0002, 0001, 0039) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0040) = 1000: ErrorCatalog(0002, 0001, 0040) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0041) = 1000: ErrorCatalog(0002, 0001, 0041) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0042) = 1000: ErrorCatalog(0002, 0001, 0042) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0043) = 1000: ErrorCatalog(0002, 0001, 0043) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0044) = 1000: ErrorCatalog(0002, 0001, 0044) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0045) = 1000: ErrorCatalog(0002, 0001, 0045) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0046) = 1000: ErrorCatalog(0002, 0001, 0046) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0047) = 1000: ErrorCatalog(0002, 0001, 0047) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0048) = 1000: ErrorCatalog(0002, 0001, 0048) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0049) = 1000: ErrorCatalog(0002, 0001, 0049) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0050) = 1000: ErrorCatalog(0002, 0001, 0050) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0051) = 1000: ErrorCatalog(0002, 0001, 0051) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0052) = 1000: ErrorCatalog(0002, 0001, 0052) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0053) = 1000: ErrorCatalog(0002, 0001, 0053) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0054) = 1000: ErrorCatalog(0002, 0001, 0054) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0055) = 1000: ErrorCatalog(0002, 0001, 0055) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0056) = 1000: ErrorCatalog(0002, 0001, 0056) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0057) = 1000: ErrorCatalog(0002, 0001, 0057) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0058) = 1000: ErrorCatalog(0002, 0001, 0058) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0059) = 1000: ErrorCatalog(0002, 0001, 0059) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0060) = 1000: ErrorCatalog(0002, 0001, 0060) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0061) = 1000: ErrorCatalog(0002, 0001, 0061) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0062) = 1000: ErrorCatalog(0002, 0001, 0062) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0063) = 1000: ErrorCatalog(0002, 0001, 0063) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0064) = 1000: ErrorCatalog(0002, 0001, 0064) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0065) = 1000: ErrorCatalog(0002, 0001, 0065) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0066) = 1000: ErrorCatalog(0002, 0001, 0066) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0067) = 1000: ErrorCatalog(0002, 0001, 0067) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0068) = 1000: ErrorCatalog(0002, 0001, 0068) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0069) = 1000: ErrorCatalog(0002, 0001, 0069) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0070) = 1000: ErrorCatalog(0002, 0001, 0070) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0071) = 1000: ErrorCatalog(0002, 0001, 0071) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0072) = 1000: ErrorCatalog(0002, 0001, 0072) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0073) = 1000: ErrorCatalog(0002, 0001, 0073) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0074) = 1000: ErrorCatalog(0002, 0001, 0074) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0075) = 1000: ErrorCatalog(0002, 0001, 0075) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0076) = 1000: ErrorCatalog(0002, 0001, 0076) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0077) = 1000: ErrorCatalog(0002, 0001, 0077) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0078) = 1000: ErrorCatalog(0002, 0001, 0078) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0079) = 1000: ErrorCatalog(0002, 0001, 0079) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0080) = 1000: ErrorCatalog(0002, 0001, 0080) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0081) = 1000: ErrorCatalog(0002, 0001, 0081) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0082) = 1000: ErrorCatalog(0002, 0001, 0082) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0083) = 1000: ErrorCatalog(0002, 0001, 0083) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0084) = 1000: ErrorCatalog(0002, 0001, 0084) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0085) = 1000: ErrorCatalog(0002, 0001, 0085) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0086) = 1000: ErrorCatalog(0002, 0001, 0086) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0087) = 1000: ErrorCatalog(0002, 0001, 0087) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0088) = 1000: ErrorCatalog(0002, 0001, 0088) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0089) = 1000: ErrorCatalog(0002, 0001, 0089) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0090) = 1000: ErrorCatalog(0002, 0001, 0090) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0091) = 1000: ErrorCatalog(0002, 0001, 0091) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0092) = 1000: ErrorCatalog(0002, 0001, 0092) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0093) = 1000: ErrorCatalog(0002, 0001, 0093) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0094) = 1000: ErrorCatalog(0002, 0001, 0094) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0095) = 1000: ErrorCatalog(0002, 0001, 0095) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0096) = 1000: ErrorCatalog(0002, 0001, 0096) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0097) = 1000: ErrorCatalog(0002, 0001, 0097) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0098) = 1000: ErrorCatalog(0002, 0001, 0098) = "PLACEHOLDER"
            ErrorCatalog(0002, 0000, 0099) = 1000: ErrorCatalog(0002, 0001, 0099) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0000) = 1000: ErrorCatalog(0003, 0001, 0000) = "Workbook doesnt exist"
            ErrorCatalog(0003, 0000, 0001) = 1000: ErrorCatalog(0003, 0001, 0001) = "Workbook already exists"
            ErrorCatalog(0003, 0000, 0002) = 1000: ErrorCatalog(0003, 0001, 0002) = "Dependency missing"
            ErrorCatalog(0003, 0000, 0003) = 1000: ErrorCatalog(0003, 0001, 0003) = "Object not Initialized"
            ErrorCatalog(0003, 0000, 0004) = 1000: ErrorCatalog(0003, 0001, 0004) = "Not available in Workbook"
            ErrorCatalog(0003, 0000, 0005) = 1000: ErrorCatalog(0003, 0001, 0005) = "Component doesnt exists"
            ErrorCatalog(0003, 0000, 0006) = 1000: ErrorCatalog(0003, 0001, 0006) = "Component already exists"
            ErrorCatalog(0003, 0000, 0007) = 1000: ErrorCatalog(0003, 0001, 0007) = "Couldnt Create Component"
            ErrorCatalog(0003, 0000, 0008) = 1000: ErrorCatalog(0003, 0001, 0008) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0009) = 1000: ErrorCatalog(0003, 0001, 0009) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0010) = 1000: ErrorCatalog(0003, 0001, 0010) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0011) = 1000: ErrorCatalog(0003, 0001, 0011) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0012) = 1000: ErrorCatalog(0003, 0001, 0012) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0013) = 1000: ErrorCatalog(0003, 0001, 0013) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0014) = 1000: ErrorCatalog(0003, 0001, 0014) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0015) = 1000: ErrorCatalog(0003, 0001, 0015) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0016) = 1000: ErrorCatalog(0003, 0001, 0016) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0017) = 1000: ErrorCatalog(0003, 0001, 0017) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0018) = 1000: ErrorCatalog(0003, 0001, 0018) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0019) = 1000: ErrorCatalog(0003, 0001, 0019) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0020) = 1000: ErrorCatalog(0003, 0001, 0020) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0021) = 1000: ErrorCatalog(0003, 0001, 0021) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0022) = 1000: ErrorCatalog(0003, 0001, 0022) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0023) = 1000: ErrorCatalog(0003, 0001, 0023) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0024) = 1000: ErrorCatalog(0003, 0001, 0024) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0025) = 1000: ErrorCatalog(0003, 0001, 0025) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0026) = 1000: ErrorCatalog(0003, 0001, 0026) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0027) = 1000: ErrorCatalog(0003, 0001, 0027) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0028) = 1000: ErrorCatalog(0003, 0001, 0028) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0029) = 1000: ErrorCatalog(0003, 0001, 0029) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0030) = 1000: ErrorCatalog(0003, 0001, 0030) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0031) = 1000: ErrorCatalog(0003, 0001, 0031) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0032) = 1000: ErrorCatalog(0003, 0001, 0032) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0033) = 1000: ErrorCatalog(0003, 0001, 0033) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0034) = 1000: ErrorCatalog(0003, 0001, 0034) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0035) = 1000: ErrorCatalog(0003, 0001, 0035) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0036) = 1000: ErrorCatalog(0003, 0001, 0036) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0037) = 1000: ErrorCatalog(0003, 0001, 0037) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0038) = 1000: ErrorCatalog(0003, 0001, 0038) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0039) = 1000: ErrorCatalog(0003, 0001, 0039) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0040) = 1000: ErrorCatalog(0003, 0001, 0040) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0041) = 1000: ErrorCatalog(0003, 0001, 0041) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0042) = 1000: ErrorCatalog(0003, 0001, 0042) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0043) = 1000: ErrorCatalog(0003, 0001, 0043) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0044) = 1000: ErrorCatalog(0003, 0001, 0044) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0045) = 1000: ErrorCatalog(0003, 0001, 0045) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0046) = 1000: ErrorCatalog(0003, 0001, 0046) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0047) = 1000: ErrorCatalog(0003, 0001, 0047) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0048) = 1000: ErrorCatalog(0003, 0001, 0048) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0049) = 1000: ErrorCatalog(0003, 0001, 0049) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0050) = 1000: ErrorCatalog(0003, 0001, 0050) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0051) = 1000: ErrorCatalog(0003, 0001, 0051) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0052) = 1000: ErrorCatalog(0003, 0001, 0052) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0053) = 1000: ErrorCatalog(0003, 0001, 0053) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0054) = 1000: ErrorCatalog(0003, 0001, 0054) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0055) = 1000: ErrorCatalog(0003, 0001, 0055) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0056) = 1000: ErrorCatalog(0003, 0001, 0056) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0057) = 1000: ErrorCatalog(0003, 0001, 0057) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0058) = 1000: ErrorCatalog(0003, 0001, 0058) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0059) = 1000: ErrorCatalog(0003, 0001, 0059) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0060) = 1000: ErrorCatalog(0003, 0001, 0060) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0061) = 1000: ErrorCatalog(0003, 0001, 0061) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0062) = 1000: ErrorCatalog(0003, 0001, 0062) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0063) = 1000: ErrorCatalog(0003, 0001, 0063) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0064) = 1000: ErrorCatalog(0003, 0001, 0064) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0065) = 1000: ErrorCatalog(0003, 0001, 0065) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0066) = 1000: ErrorCatalog(0003, 0001, 0066) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0067) = 1000: ErrorCatalog(0003, 0001, 0067) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0068) = 1000: ErrorCatalog(0003, 0001, 0068) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0069) = 1000: ErrorCatalog(0003, 0001, 0069) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0070) = 1000: ErrorCatalog(0003, 0001, 0070) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0071) = 1000: ErrorCatalog(0003, 0001, 0071) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0072) = 1000: ErrorCatalog(0003, 0001, 0072) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0073) = 1000: ErrorCatalog(0003, 0001, 0073) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0074) = 1000: ErrorCatalog(0003, 0001, 0074) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0075) = 1000: ErrorCatalog(0003, 0001, 0075) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0076) = 1000: ErrorCatalog(0003, 0001, 0076) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0077) = 1000: ErrorCatalog(0003, 0001, 0077) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0078) = 1000: ErrorCatalog(0003, 0001, 0078) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0079) = 1000: ErrorCatalog(0003, 0001, 0079) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0080) = 1000: ErrorCatalog(0003, 0001, 0080) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0081) = 1000: ErrorCatalog(0003, 0001, 0081) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0082) = 1000: ErrorCatalog(0003, 0001, 0082) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0083) = 1000: ErrorCatalog(0003, 0001, 0083) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0084) = 1000: ErrorCatalog(0003, 0001, 0084) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0085) = 1000: ErrorCatalog(0003, 0001, 0085) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0086) = 1000: ErrorCatalog(0003, 0001, 0086) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0087) = 1000: ErrorCatalog(0003, 0001, 0087) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0088) = 1000: ErrorCatalog(0003, 0001, 0088) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0089) = 1000: ErrorCatalog(0003, 0001, 0089) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0090) = 1000: ErrorCatalog(0003, 0001, 0090) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0091) = 1000: ErrorCatalog(0003, 0001, 0091) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0092) = 1000: ErrorCatalog(0003, 0001, 0092) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0093) = 1000: ErrorCatalog(0003, 0001, 0093) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0094) = 1000: ErrorCatalog(0003, 0001, 0094) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0095) = 1000: ErrorCatalog(0003, 0001, 0095) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0096) = 1000: ErrorCatalog(0003, 0001, 0096) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0097) = 1000: ErrorCatalog(0003, 0001, 0097) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0098) = 1000: ErrorCatalog(0003, 0001, 0098) = "PLACEHOLDER"
            ErrorCatalog(0003, 0000, 0099) = 1000: ErrorCatalog(0003, 0001, 0099) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0000) = 1000: ErrorCatalog(0004, 0001, 0000) = "Worksheet already exists"
            ErrorCatalog(0004, 0000, 0001) = 1000: ErrorCatalog(0004, 0001, 0001) = "Worksheet doesnt exist"
            ErrorCatalog(0004, 0000, 0002) = 1000: ErrorCatalog(0004, 0001, 0002) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0003) = 1000: ErrorCatalog(0004, 0001, 0003) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0004) = 1000: ErrorCatalog(0004, 0001, 0004) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0005) = 1000: ErrorCatalog(0004, 0001, 0005) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0006) = 1000: ErrorCatalog(0004, 0001, 0006) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0007) = 1000: ErrorCatalog(0004, 0001, 0007) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0008) = 1000: ErrorCatalog(0004, 0001, 0008) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0009) = 1000: ErrorCatalog(0004, 0001, 0009) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0010) = 1000: ErrorCatalog(0004, 0001, 0010) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0011) = 1000: ErrorCatalog(0004, 0001, 0011) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0012) = 1000: ErrorCatalog(0004, 0001, 0012) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0013) = 1000: ErrorCatalog(0004, 0001, 0013) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0014) = 1000: ErrorCatalog(0004, 0001, 0014) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0015) = 1000: ErrorCatalog(0004, 0001, 0015) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0016) = 1000: ErrorCatalog(0004, 0001, 0016) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0017) = 1000: ErrorCatalog(0004, 0001, 0017) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0018) = 1000: ErrorCatalog(0004, 0001, 0018) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0019) = 1000: ErrorCatalog(0004, 0001, 0019) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0020) = 1000: ErrorCatalog(0004, 0001, 0020) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0021) = 1000: ErrorCatalog(0004, 0001, 0021) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0022) = 1000: ErrorCatalog(0004, 0001, 0022) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0023) = 1000: ErrorCatalog(0004, 0001, 0023) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0024) = 1000: ErrorCatalog(0004, 0001, 0024) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0025) = 1000: ErrorCatalog(0004, 0001, 0025) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0026) = 1000: ErrorCatalog(0004, 0001, 0026) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0027) = 1000: ErrorCatalog(0004, 0001, 0027) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0028) = 1000: ErrorCatalog(0004, 0001, 0028) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0029) = 1000: ErrorCatalog(0004, 0001, 0029) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0030) = 1000: ErrorCatalog(0004, 0001, 0030) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0031) = 1000: ErrorCatalog(0004, 0001, 0031) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0032) = 1000: ErrorCatalog(0004, 0001, 0032) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0033) = 1000: ErrorCatalog(0004, 0001, 0033) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0034) = 1000: ErrorCatalog(0004, 0001, 0034) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0035) = 1000: ErrorCatalog(0004, 0001, 0035) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0036) = 1000: ErrorCatalog(0004, 0001, 0036) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0037) = 1000: ErrorCatalog(0004, 0001, 0037) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0038) = 1000: ErrorCatalog(0004, 0001, 0038) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0039) = 1000: ErrorCatalog(0004, 0001, 0039) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0040) = 1000: ErrorCatalog(0004, 0001, 0040) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0041) = 1000: ErrorCatalog(0004, 0001, 0041) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0042) = 1000: ErrorCatalog(0004, 0001, 0042) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0043) = 1000: ErrorCatalog(0004, 0001, 0043) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0044) = 1000: ErrorCatalog(0004, 0001, 0044) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0045) = 1000: ErrorCatalog(0004, 0001, 0045) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0046) = 1000: ErrorCatalog(0004, 0001, 0046) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0047) = 1000: ErrorCatalog(0004, 0001, 0047) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0048) = 1000: ErrorCatalog(0004, 0001, 0048) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0049) = 1000: ErrorCatalog(0004, 0001, 0049) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0050) = 1000: ErrorCatalog(0004, 0001, 0050) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0051) = 1000: ErrorCatalog(0004, 0001, 0051) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0052) = 1000: ErrorCatalog(0004, 0001, 0052) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0053) = 1000: ErrorCatalog(0004, 0001, 0053) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0054) = 1000: ErrorCatalog(0004, 0001, 0054) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0055) = 1000: ErrorCatalog(0004, 0001, 0055) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0056) = 1000: ErrorCatalog(0004, 0001, 0056) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0057) = 1000: ErrorCatalog(0004, 0001, 0057) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0058) = 1000: ErrorCatalog(0004, 0001, 0058) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0059) = 1000: ErrorCatalog(0004, 0001, 0059) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0060) = 1000: ErrorCatalog(0004, 0001, 0060) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0061) = 1000: ErrorCatalog(0004, 0001, 0061) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0062) = 1000: ErrorCatalog(0004, 0001, 0062) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0063) = 1000: ErrorCatalog(0004, 0001, 0063) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0064) = 1000: ErrorCatalog(0004, 0001, 0064) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0065) = 1000: ErrorCatalog(0004, 0001, 0065) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0066) = 1000: ErrorCatalog(0004, 0001, 0066) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0067) = 1000: ErrorCatalog(0004, 0001, 0067) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0068) = 1000: ErrorCatalog(0004, 0001, 0068) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0069) = 1000: ErrorCatalog(0004, 0001, 0069) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0070) = 1000: ErrorCatalog(0004, 0001, 0070) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0071) = 1000: ErrorCatalog(0004, 0001, 0071) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0072) = 1000: ErrorCatalog(0004, 0001, 0072) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0073) = 1000: ErrorCatalog(0004, 0001, 0073) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0074) = 1000: ErrorCatalog(0004, 0001, 0074) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0075) = 1000: ErrorCatalog(0004, 0001, 0075) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0076) = 1000: ErrorCatalog(0004, 0001, 0076) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0077) = 1000: ErrorCatalog(0004, 0001, 0077) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0078) = 1000: ErrorCatalog(0004, 0001, 0078) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0079) = 1000: ErrorCatalog(0004, 0001, 0079) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0080) = 1000: ErrorCatalog(0004, 0001, 0080) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0081) = 1000: ErrorCatalog(0004, 0001, 0081) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0082) = 1000: ErrorCatalog(0004, 0001, 0082) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0083) = 1000: ErrorCatalog(0004, 0001, 0083) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0084) = 1000: ErrorCatalog(0004, 0001, 0084) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0085) = 1000: ErrorCatalog(0004, 0001, 0085) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0086) = 1000: ErrorCatalog(0004, 0001, 0086) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0087) = 1000: ErrorCatalog(0004, 0001, 0087) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0088) = 1000: ErrorCatalog(0004, 0001, 0088) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0089) = 1000: ErrorCatalog(0004, 0001, 0089) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0090) = 1000: ErrorCatalog(0004, 0001, 0090) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0091) = 1000: ErrorCatalog(0004, 0001, 0091) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0092) = 1000: ErrorCatalog(0004, 0001, 0092) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0093) = 1000: ErrorCatalog(0004, 0001, 0093) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0094) = 1000: ErrorCatalog(0004, 0001, 0094) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0095) = 1000: ErrorCatalog(0004, 0001, 0095) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0096) = 1000: ErrorCatalog(0004, 0001, 0096) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0097) = 1000: ErrorCatalog(0004, 0001, 0097) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0098) = 1000: ErrorCatalog(0004, 0001, 0098) = "PLACEHOLDER"
            ErrorCatalog(0004, 0000, 0099) = 1000: ErrorCatalog(0004, 0001, 0099) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0000) = 1000: ErrorCatalog(0005, 0001, 0000) = "Invalid Value"
            ErrorCatalog(0005, 0000, 0001) = 1000: ErrorCatalog(0005, 0001, 0001) = "Value is Nothing"
            ErrorCatalog(0005, 0000, 0002) = 1000: ErrorCatalog(0005, 0001, 0002) = "Value Underflow"
            ErrorCatalog(0005, 0000, 0003) = 1000: ErrorCatalog(0005, 0001, 0003) = "Value Overflow"
            ErrorCatalog(0005, 0000, 0004) = 1000: ErrorCatalog(0005, 0001, 0004) = "Object not Initialized"
            ErrorCatalog(0005, 0000, 0005) = 1000: ErrorCatalog(0005, 0001, 0005) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0006) = 1000: ErrorCatalog(0005, 0001, 0006) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0007) = 1000: ErrorCatalog(0005, 0001, 0007) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0008) = 1000: ErrorCatalog(0005, 0001, 0008) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0009) = 1000: ErrorCatalog(0005, 0001, 0009) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0010) = 1000: ErrorCatalog(0005, 0001, 0010) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0011) = 1000: ErrorCatalog(0005, 0001, 0011) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0012) = 1000: ErrorCatalog(0005, 0001, 0012) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0013) = 1000: ErrorCatalog(0005, 0001, 0013) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0014) = 1000: ErrorCatalog(0005, 0001, 0014) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0015) = 1000: ErrorCatalog(0005, 0001, 0015) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0016) = 1000: ErrorCatalog(0005, 0001, 0016) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0017) = 1000: ErrorCatalog(0005, 0001, 0017) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0018) = 1000: ErrorCatalog(0005, 0001, 0018) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0019) = 1000: ErrorCatalog(0005, 0001, 0019) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0020) = 1000: ErrorCatalog(0005, 0001, 0020) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0021) = 1000: ErrorCatalog(0005, 0001, 0021) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0022) = 1000: ErrorCatalog(0005, 0001, 0022) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0023) = 1000: ErrorCatalog(0005, 0001, 0023) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0024) = 1000: ErrorCatalog(0005, 0001, 0024) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0025) = 1000: ErrorCatalog(0005, 0001, 0025) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0026) = 1000: ErrorCatalog(0005, 0001, 0026) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0027) = 1000: ErrorCatalog(0005, 0001, 0027) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0028) = 1000: ErrorCatalog(0005, 0001, 0028) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0029) = 1000: ErrorCatalog(0005, 0001, 0029) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0030) = 1000: ErrorCatalog(0005, 0001, 0030) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0031) = 1000: ErrorCatalog(0005, 0001, 0031) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0032) = 1000: ErrorCatalog(0005, 0001, 0032) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0033) = 1000: ErrorCatalog(0005, 0001, 0033) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0034) = 1000: ErrorCatalog(0005, 0001, 0034) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0035) = 1000: ErrorCatalog(0005, 0001, 0035) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0036) = 1000: ErrorCatalog(0005, 0001, 0036) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0037) = 1000: ErrorCatalog(0005, 0001, 0037) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0038) = 1000: ErrorCatalog(0005, 0001, 0038) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0039) = 1000: ErrorCatalog(0005, 0001, 0039) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0040) = 1000: ErrorCatalog(0005, 0001, 0040) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0041) = 1000: ErrorCatalog(0005, 0001, 0041) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0042) = 1000: ErrorCatalog(0005, 0001, 0042) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0043) = 1000: ErrorCatalog(0005, 0001, 0043) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0044) = 1000: ErrorCatalog(0005, 0001, 0044) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0045) = 1000: ErrorCatalog(0005, 0001, 0045) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0046) = 1000: ErrorCatalog(0005, 0001, 0046) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0047) = 1000: ErrorCatalog(0005, 0001, 0047) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0048) = 1000: ErrorCatalog(0005, 0001, 0048) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0049) = 1000: ErrorCatalog(0005, 0001, 0049) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0050) = 1000: ErrorCatalog(0005, 0001, 0050) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0051) = 1000: ErrorCatalog(0005, 0001, 0051) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0052) = 1000: ErrorCatalog(0005, 0001, 0052) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0053) = 1000: ErrorCatalog(0005, 0001, 0053) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0054) = 1000: ErrorCatalog(0005, 0001, 0054) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0055) = 1000: ErrorCatalog(0005, 0001, 0055) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0056) = 1000: ErrorCatalog(0005, 0001, 0056) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0057) = 1000: ErrorCatalog(0005, 0001, 0057) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0058) = 1000: ErrorCatalog(0005, 0001, 0058) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0059) = 1000: ErrorCatalog(0005, 0001, 0059) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0060) = 1000: ErrorCatalog(0005, 0001, 0060) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0061) = 1000: ErrorCatalog(0005, 0001, 0061) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0062) = 1000: ErrorCatalog(0005, 0001, 0062) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0063) = 1000: ErrorCatalog(0005, 0001, 0063) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0064) = 1000: ErrorCatalog(0005, 0001, 0064) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0065) = 1000: ErrorCatalog(0005, 0001, 0065) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0066) = 1000: ErrorCatalog(0005, 0001, 0066) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0067) = 1000: ErrorCatalog(0005, 0001, 0067) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0068) = 1000: ErrorCatalog(0005, 0001, 0068) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0069) = 1000: ErrorCatalog(0005, 0001, 0069) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0070) = 1000: ErrorCatalog(0005, 0001, 0070) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0071) = 1000: ErrorCatalog(0005, 0001, 0071) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0072) = 1000: ErrorCatalog(0005, 0001, 0072) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0073) = 1000: ErrorCatalog(0005, 0001, 0073) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0074) = 1000: ErrorCatalog(0005, 0001, 0074) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0075) = 1000: ErrorCatalog(0005, 0001, 0075) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0076) = 1000: ErrorCatalog(0005, 0001, 0076) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0077) = 1000: ErrorCatalog(0005, 0001, 0077) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0078) = 1000: ErrorCatalog(0005, 0001, 0078) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0079) = 1000: ErrorCatalog(0005, 0001, 0079) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0080) = 1000: ErrorCatalog(0005, 0001, 0080) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0081) = 1000: ErrorCatalog(0005, 0001, 0081) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0082) = 1000: ErrorCatalog(0005, 0001, 0082) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0083) = 1000: ErrorCatalog(0005, 0001, 0083) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0084) = 1000: ErrorCatalog(0005, 0001, 0084) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0085) = 1000: ErrorCatalog(0005, 0001, 0085) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0086) = 1000: ErrorCatalog(0005, 0001, 0086) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0087) = 1000: ErrorCatalog(0005, 0001, 0087) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0088) = 1000: ErrorCatalog(0005, 0001, 0088) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0089) = 1000: ErrorCatalog(0005, 0001, 0089) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0090) = 1000: ErrorCatalog(0005, 0001, 0090) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0091) = 1000: ErrorCatalog(0005, 0001, 0091) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0092) = 1000: ErrorCatalog(0005, 0001, 0092) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0093) = 1000: ErrorCatalog(0005, 0001, 0093) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0094) = 1000: ErrorCatalog(0005, 0001, 0094) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0095) = 1000: ErrorCatalog(0005, 0001, 0095) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0096) = 1000: ErrorCatalog(0005, 0001, 0096) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0097) = 1000: ErrorCatalog(0005, 0001, 0097) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0098) = 1000: ErrorCatalog(0005, 0001, 0098) = "PLACEHOLDER"
            ErrorCatalog(0005, 0000, 0099) = 1000: ErrorCatalog(0005, 0001, 0099) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0000) = 1000: ErrorCatalog(0006, 0001, 0000) = "Invalid Value"
            ErrorCatalog(0006, 0000, 0001) = 1000: ErrorCatalog(0006, 0001, 0001) = "Value is Nothing"
            ErrorCatalog(0006, 0000, 0002) = 1000: ErrorCatalog(0006, 0001, 0002) = "Value Underflow"
            ErrorCatalog(0006, 0000, 0003) = 1000: ErrorCatalog(0006, 0001, 0003) = "Value Overflow"
            ErrorCatalog(0006, 0000, 0004) = 1000: ErrorCatalog(0006, 0001, 0004) = "Object not Initialized"
            ErrorCatalog(0006, 0000, 0005) = 1000: ErrorCatalog(0006, 0001, 0005) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0006) = 1000: ErrorCatalog(0006, 0001, 0006) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0007) = 1000: ErrorCatalog(0006, 0001, 0007) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0008) = 1000: ErrorCatalog(0006, 0001, 0008) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0009) = 1000: ErrorCatalog(0006, 0001, 0009) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0010) = 1000: ErrorCatalog(0006, 0001, 0010) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0011) = 1000: ErrorCatalog(0006, 0001, 0011) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0012) = 1000: ErrorCatalog(0006, 0001, 0012) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0013) = 1000: ErrorCatalog(0006, 0001, 0013) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0014) = 1000: ErrorCatalog(0006, 0001, 0014) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0015) = 1000: ErrorCatalog(0006, 0001, 0015) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0016) = 1000: ErrorCatalog(0006, 0001, 0016) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0017) = 1000: ErrorCatalog(0006, 0001, 0017) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0018) = 1000: ErrorCatalog(0006, 0001, 0018) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0019) = 1000: ErrorCatalog(0006, 0001, 0019) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0020) = 1000: ErrorCatalog(0006, 0001, 0020) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0021) = 1000: ErrorCatalog(0006, 0001, 0021) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0022) = 1000: ErrorCatalog(0006, 0001, 0022) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0023) = 1000: ErrorCatalog(0006, 0001, 0023) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0024) = 1000: ErrorCatalog(0006, 0001, 0024) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0025) = 1000: ErrorCatalog(0006, 0001, 0025) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0026) = 1000: ErrorCatalog(0006, 0001, 0026) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0027) = 1000: ErrorCatalog(0006, 0001, 0027) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0028) = 1000: ErrorCatalog(0006, 0001, 0028) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0029) = 1000: ErrorCatalog(0006, 0001, 0029) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0030) = 1000: ErrorCatalog(0006, 0001, 0030) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0031) = 1000: ErrorCatalog(0006, 0001, 0031) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0032) = 1000: ErrorCatalog(0006, 0001, 0032) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0033) = 1000: ErrorCatalog(0006, 0001, 0033) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0034) = 1000: ErrorCatalog(0006, 0001, 0034) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0035) = 1000: ErrorCatalog(0006, 0001, 0035) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0036) = 1000: ErrorCatalog(0006, 0001, 0036) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0037) = 1000: ErrorCatalog(0006, 0001, 0037) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0038) = 1000: ErrorCatalog(0006, 0001, 0038) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0039) = 1000: ErrorCatalog(0006, 0001, 0039) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0040) = 1000: ErrorCatalog(0006, 0001, 0040) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0041) = 1000: ErrorCatalog(0006, 0001, 0041) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0042) = 1000: ErrorCatalog(0006, 0001, 0042) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0043) = 1000: ErrorCatalog(0006, 0001, 0043) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0044) = 1000: ErrorCatalog(0006, 0001, 0044) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0045) = 1000: ErrorCatalog(0006, 0001, 0045) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0046) = 1000: ErrorCatalog(0006, 0001, 0046) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0047) = 1000: ErrorCatalog(0006, 0001, 0047) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0048) = 1000: ErrorCatalog(0006, 0001, 0048) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0049) = 1000: ErrorCatalog(0006, 0001, 0049) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0050) = 1000: ErrorCatalog(0006, 0001, 0050) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0051) = 1000: ErrorCatalog(0006, 0001, 0051) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0052) = 1000: ErrorCatalog(0006, 0001, 0052) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0053) = 1000: ErrorCatalog(0006, 0001, 0053) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0054) = 1000: ErrorCatalog(0006, 0001, 0054) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0055) = 1000: ErrorCatalog(0006, 0001, 0055) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0056) = 1000: ErrorCatalog(0006, 0001, 0056) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0057) = 1000: ErrorCatalog(0006, 0001, 0057) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0058) = 1000: ErrorCatalog(0006, 0001, 0058) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0059) = 1000: ErrorCatalog(0006, 0001, 0059) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0060) = 1000: ErrorCatalog(0006, 0001, 0060) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0061) = 1000: ErrorCatalog(0006, 0001, 0061) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0062) = 1000: ErrorCatalog(0006, 0001, 0062) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0063) = 1000: ErrorCatalog(0006, 0001, 0063) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0064) = 1000: ErrorCatalog(0006, 0001, 0064) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0065) = 1000: ErrorCatalog(0006, 0001, 0065) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0066) = 1000: ErrorCatalog(0006, 0001, 0066) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0067) = 1000: ErrorCatalog(0006, 0001, 0067) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0068) = 1000: ErrorCatalog(0006, 0001, 0068) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0069) = 1000: ErrorCatalog(0006, 0001, 0069) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0070) = 1000: ErrorCatalog(0006, 0001, 0070) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0071) = 1000: ErrorCatalog(0006, 0001, 0071) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0072) = 1000: ErrorCatalog(0006, 0001, 0072) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0073) = 1000: ErrorCatalog(0006, 0001, 0073) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0074) = 1000: ErrorCatalog(0006, 0001, 0074) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0075) = 1000: ErrorCatalog(0006, 0001, 0075) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0076) = 1000: ErrorCatalog(0006, 0001, 0076) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0077) = 1000: ErrorCatalog(0006, 0001, 0077) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0078) = 1000: ErrorCatalog(0006, 0001, 0078) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0079) = 1000: ErrorCatalog(0006, 0001, 0079) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0080) = 1000: ErrorCatalog(0006, 0001, 0080) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0081) = 1000: ErrorCatalog(0006, 0001, 0081) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0082) = 1000: ErrorCatalog(0006, 0001, 0082) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0083) = 1000: ErrorCatalog(0006, 0001, 0083) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0084) = 1000: ErrorCatalog(0006, 0001, 0084) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0085) = 1000: ErrorCatalog(0006, 0001, 0085) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0086) = 1000: ErrorCatalog(0006, 0001, 0086) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0087) = 1000: ErrorCatalog(0006, 0001, 0087) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0088) = 1000: ErrorCatalog(0006, 0001, 0088) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0089) = 1000: ErrorCatalog(0006, 0001, 0089) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0090) = 1000: ErrorCatalog(0006, 0001, 0090) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0091) = 1000: ErrorCatalog(0006, 0001, 0091) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0092) = 1000: ErrorCatalog(0006, 0001, 0092) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0093) = 1000: ErrorCatalog(0006, 0001, 0093) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0094) = 1000: ErrorCatalog(0006, 0001, 0094) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0095) = 1000: ErrorCatalog(0006, 0001, 0095) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0096) = 1000: ErrorCatalog(0006, 0001, 0096) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0097) = 1000: ErrorCatalog(0006, 0001, 0097) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0098) = 1000: ErrorCatalog(0006, 0001, 0098) = "PLACEHOLDER"
            ErrorCatalog(0006, 0000, 0099) = 1000: ErrorCatalog(0006, 0001, 0099) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0000) = 1000: ErrorCatalog(0007, 0001, 0000) = "Invalid Value"
            ErrorCatalog(0007, 0000, 0001) = 1000: ErrorCatalog(0007, 0001, 0001) = "Value is Nothing"
            ErrorCatalog(0007, 0000, 0002) = 1000: ErrorCatalog(0007, 0001, 0002) = "Value Underflow"
            ErrorCatalog(0007, 0000, 0003) = 1000: ErrorCatalog(0007, 0001, 0003) = "Value Overflow"
            ErrorCatalog(0007, 0000, 0004) = 1000: ErrorCatalog(0007, 0001, 0004) = "Object not Initialized"
            ErrorCatalog(0007, 0000, 0005) = 1000: ErrorCatalog(0007, 0001, 0005) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0006) = 1000: ErrorCatalog(0007, 0001, 0006) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0007) = 1000: ErrorCatalog(0007, 0001, 0007) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0008) = 1000: ErrorCatalog(0007, 0001, 0008) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0009) = 1000: ErrorCatalog(0007, 0001, 0009) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0010) = 1000: ErrorCatalog(0007, 0001, 0010) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0011) = 1000: ErrorCatalog(0007, 0001, 0011) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0012) = 1000: ErrorCatalog(0007, 0001, 0012) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0013) = 1000: ErrorCatalog(0007, 0001, 0013) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0014) = 1000: ErrorCatalog(0007, 0001, 0014) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0015) = 1000: ErrorCatalog(0007, 0001, 0015) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0016) = 1000: ErrorCatalog(0007, 0001, 0016) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0017) = 1000: ErrorCatalog(0007, 0001, 0017) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0018) = 1000: ErrorCatalog(0007, 0001, 0018) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0019) = 1000: ErrorCatalog(0007, 0001, 0019) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0020) = 1000: ErrorCatalog(0007, 0001, 0020) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0021) = 1000: ErrorCatalog(0007, 0001, 0021) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0022) = 1000: ErrorCatalog(0007, 0001, 0022) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0023) = 1000: ErrorCatalog(0007, 0001, 0023) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0024) = 1000: ErrorCatalog(0007, 0001, 0024) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0025) = 1000: ErrorCatalog(0007, 0001, 0025) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0026) = 1000: ErrorCatalog(0007, 0001, 0026) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0027) = 1000: ErrorCatalog(0007, 0001, 0027) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0028) = 1000: ErrorCatalog(0007, 0001, 0028) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0029) = 1000: ErrorCatalog(0007, 0001, 0029) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0030) = 1000: ErrorCatalog(0007, 0001, 0030) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0031) = 1000: ErrorCatalog(0007, 0001, 0031) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0032) = 1000: ErrorCatalog(0007, 0001, 0032) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0033) = 1000: ErrorCatalog(0007, 0001, 0033) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0034) = 1000: ErrorCatalog(0007, 0001, 0034) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0035) = 1000: ErrorCatalog(0007, 0001, 0035) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0036) = 1000: ErrorCatalog(0007, 0001, 0036) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0037) = 1000: ErrorCatalog(0007, 0001, 0037) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0038) = 1000: ErrorCatalog(0007, 0001, 0038) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0039) = 1000: ErrorCatalog(0007, 0001, 0039) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0040) = 1000: ErrorCatalog(0007, 0001, 0040) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0041) = 1000: ErrorCatalog(0007, 0001, 0041) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0042) = 1000: ErrorCatalog(0007, 0001, 0042) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0043) = 1000: ErrorCatalog(0007, 0001, 0043) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0044) = 1000: ErrorCatalog(0007, 0001, 0044) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0045) = 1000: ErrorCatalog(0007, 0001, 0045) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0046) = 1000: ErrorCatalog(0007, 0001, 0046) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0047) = 1000: ErrorCatalog(0007, 0001, 0047) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0048) = 1000: ErrorCatalog(0007, 0001, 0048) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0049) = 1000: ErrorCatalog(0007, 0001, 0049) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0050) = 1000: ErrorCatalog(0007, 0001, 0050) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0051) = 1000: ErrorCatalog(0007, 0001, 0051) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0052) = 1000: ErrorCatalog(0007, 0001, 0052) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0053) = 1000: ErrorCatalog(0007, 0001, 0053) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0054) = 1000: ErrorCatalog(0007, 0001, 0054) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0055) = 1000: ErrorCatalog(0007, 0001, 0055) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0056) = 1000: ErrorCatalog(0007, 0001, 0056) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0057) = 1000: ErrorCatalog(0007, 0001, 0057) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0058) = 1000: ErrorCatalog(0007, 0001, 0058) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0059) = 1000: ErrorCatalog(0007, 0001, 0059) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0060) = 1000: ErrorCatalog(0007, 0001, 0060) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0061) = 1000: ErrorCatalog(0007, 0001, 0061) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0062) = 1000: ErrorCatalog(0007, 0001, 0062) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0063) = 1000: ErrorCatalog(0007, 0001, 0063) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0064) = 1000: ErrorCatalog(0007, 0001, 0064) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0065) = 1000: ErrorCatalog(0007, 0001, 0065) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0066) = 1000: ErrorCatalog(0007, 0001, 0066) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0067) = 1000: ErrorCatalog(0007, 0001, 0067) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0068) = 1000: ErrorCatalog(0007, 0001, 0068) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0069) = 1000: ErrorCatalog(0007, 0001, 0069) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0070) = 1000: ErrorCatalog(0007, 0001, 0070) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0071) = 1000: ErrorCatalog(0007, 0001, 0071) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0072) = 1000: ErrorCatalog(0007, 0001, 0072) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0073) = 1000: ErrorCatalog(0007, 0001, 0073) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0074) = 1000: ErrorCatalog(0007, 0001, 0074) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0075) = 1000: ErrorCatalog(0007, 0001, 0075) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0076) = 1000: ErrorCatalog(0007, 0001, 0076) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0077) = 1000: ErrorCatalog(0007, 0001, 0077) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0078) = 1000: ErrorCatalog(0007, 0001, 0078) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0079) = 1000: ErrorCatalog(0007, 0001, 0079) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0080) = 1000: ErrorCatalog(0007, 0001, 0080) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0081) = 1000: ErrorCatalog(0007, 0001, 0081) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0082) = 1000: ErrorCatalog(0007, 0001, 0082) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0083) = 1000: ErrorCatalog(0007, 0001, 0083) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0084) = 1000: ErrorCatalog(0007, 0001, 0084) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0085) = 1000: ErrorCatalog(0007, 0001, 0085) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0086) = 1000: ErrorCatalog(0007, 0001, 0086) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0087) = 1000: ErrorCatalog(0007, 0001, 0087) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0088) = 1000: ErrorCatalog(0007, 0001, 0088) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0089) = 1000: ErrorCatalog(0007, 0001, 0089) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0090) = 1000: ErrorCatalog(0007, 0001, 0090) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0091) = 1000: ErrorCatalog(0007, 0001, 0091) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0092) = 1000: ErrorCatalog(0007, 0001, 0092) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0093) = 1000: ErrorCatalog(0007, 0001, 0093) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0094) = 1000: ErrorCatalog(0007, 0001, 0094) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0095) = 1000: ErrorCatalog(0007, 0001, 0095) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0096) = 1000: ErrorCatalog(0007, 0001, 0096) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0097) = 1000: ErrorCatalog(0007, 0001, 0097) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0098) = 1000: ErrorCatalog(0007, 0001, 0098) = "PLACEHOLDER"
            ErrorCatalog(0007, 0000, 0099) = 1000: ErrorCatalog(0007, 0001, 0099) = "PLACEHOLDER"
        Initialized = True
        End If

    End Sub
'
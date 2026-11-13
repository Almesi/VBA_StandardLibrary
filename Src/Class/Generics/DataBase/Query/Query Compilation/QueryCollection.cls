
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mDefinitions As New Collection
Private mRows As New Collection
Private mResolving As New Collection
Private mResolved As New Collection


'===============================================================================
' CREATION
'===============================================================================

Public Function Create() As QueryCollection
    Set Create = New QueryCollection
    Call Create.Initialize
End Function

Friend Sub Initialize()
    Call PreDefineBasics
End Sub

Public Function Exists(ByVal nName As String) As Boolean
    Dim Definition As QueryDefinition

    On Error Resume Next
    Set Definition = mDefinitions(nName)
    Exists = Not Definition Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

Public Function GetDefinition(ByVal nName As String) As QueryDefinition
    If Exists(nName) Then
        Set GetDefinition = mDefinitions(nName)
    Else
        Err.Raise vbObjectError + 1000, "QueryCollection", "Unknown query definition: " & nName
    End If
End Function

Private Function IsTableDefinition(ByVal nName As String) As Boolean
    IsTableDefinition = IsInCollection(mRows, nName)
End Function


'===============================================================================
' TABLE LOADING
'===============================================================================

Public Function CreateFromDataSet(ByVal nDataSet As DataSet, ByVal NameField As Field) As QueryCollection
    Set CreateFromDataSet = Create()
    Call CreateFromDataSet.AddFromDataSet(nDataSet, NameField)
End Function

Friend Sub AddFromDataSet(ByVal nDataSet As DataSet, ByVal NameField As Field)
    Dim Item As Variant
    Dim Definition As QueryDefinition
    Dim Row As RowEntry

    ' First pass: register all names and retain their rows.
    For Each Item In NameField.DataEntries

        If Exists(CStr(Item.Value)) Then
            Err.Raise vbObjectError + 1001, "QueryCollection", "Duplicate query definition: " & CStr(Item.Value)
        End If

        Set Definition = QueryDefinition.Create(CStr(Item.Value), Nothing, 0)
        Call mDefinitions.Add(Definition, Definition.Name)

        Set Row = nDataSet.RowFromEntry(Item)
        Call mRows.Add(Row, Definition.Name)

    Next Item

    ' Second pass: resolve all definitions.
    For Each Item In NameField.DataEntries
        Call ResolveDefinition(CStr(Item.Value))
    Next Item
End Sub


'===============================================================================
' DEFINITION RESOLUTION
'===============================================================================

Private Sub ResolveDefinition(ByVal nName As String)
    Dim Definition As QueryDefinition
    Dim Row As RowEntry
    Dim ErrorNumber As Long
    Dim ErrorSource As String
    Dim ErrorDescription As String

    If IsInCollection(mResolved, nName) Then Exit Sub

    If IsInCollection(mResolving, nName) Then
        Err.Raise vbObjectError + 1002, "QueryCollection", "Circular query definition detected: " & nName
    End If

    Set Definition = GetDefinition(nName)

    Call mResolving.Add(True, nName)

    On Error GoTo Failed

    Set Row = mRows(nName)
    Call DefineFromRow(Definition, Row)

    Call mResolved.Add(True, nName)
    Call mResolving.Remove(nName)

    Exit Sub

    Failed:
    ErrorNumber = Err.Number
    ErrorSource = Err.Source
    ErrorDescription = Err.Description

    On Error Resume Next
    Call mResolving.Remove(nName)
    On Error GoTo 0

    Err.Raise ErrorNumber, ErrorSource, ErrorDescription
End Sub


Private Sub DefineFromRow(ByVal Definition As QueryDefinition, ByVal nRow As RowEntry)
    Dim Entries As Collection
    Dim ColumnIndex As Long
    Dim Unexpected As String

    Set Entries = nRow.FieldEntries

    ' Column 1 contains the ID.
    ' Column 2 contains the query name.
    ' Column 3 contains the underlying command.
    ColumnIndex = 3

    Call SkipEmptyEntries(Entries)
    Call DefineCommand(Definition, Entries, ColumnIndex)
    Dim Links As Long
    Links = Definition.ArgumentLinkCount()
    If Links <> Definition.ArgumentCount Then Call Definition.SetArgumentCount(Links)

    If ColumnIndex <= Entries.Count Then
        Unexpected = ConsumeToken(Entries, ColumnIndex, "unexpected trailing token in " & Definition.Name)

        Err.Raise vbObjectError + 1011, "QueryCollection", "Unexpected trailing token in " & Definition.Name & ": " & Unexpected
    End If
End Sub


'===============================================================================
' COMMAND PARSING
'===============================================================================

Private Sub DefineCommand(ByVal Destination As QueryDefinition, ByVal Entries As Collection, ByRef ColumnIndex As Long)
    Dim CommandName As String

    CommandName = ConsumeToken(Entries, ColumnIndex, "command name in " & Destination.Name)

    Call DefineCommandBody(Destination, CommandName, Entries, ColumnIndex)
End Sub


Private Sub DefineCommandBody(ByVal Destination As QueryDefinition, ByVal CommandName As String, ByVal Entries As Collection, ByRef ColumnIndex As Long)
    Dim ContextQuery As QueryDefinition
    Dim ArgumentIndex As Long
    Dim ExpectedCount As Long

    If Not Exists(CommandName) Then
        Err.Raise vbObjectError + 1004, "QueryCollection", "Unknown query command: " & CommandName
    End If

    Set ContextQuery = GetDefinition(CommandName)

    Call Destination.SetReturnType(ContextQuery.ReturnType)
    Call Destination.SetBasicCommand(ContextQuery.BasicCommand)

    ExpectedCount = ContextQuery.ArgumentCount

    If ExpectedCount = -1 Then

        Do While ColumnIndex <= Entries.Count
            ArgumentIndex = ArgumentIndex + 1

            Call DefineArgument(Destination, ContextQuery, Entries, ColumnIndex, ArgumentIndex)
        Loop

        ' A variable-argument command returns the type of its final child.
        If Destination.ArgumentCount > 0 Then
            Call Destination.SetReturnType(Destination.Argument(Destination.ArgumentCount).ArgumentReturn)
        End If

    Else

        For ArgumentIndex = 1 To ExpectedCount
            Call DefineArgument(Destination, ContextQuery, Entries, ColumnIndex, ArgumentIndex)
        Next ArgumentIndex

    End If
End Sub


Private Sub DefineArgument(ByVal Destination As QueryDefinition, ByVal ContextQuery As QueryDefinition, ByVal Entries As Collection, ByRef ColumnIndex As Long, ByVal OutputIndex As Long)
    Dim Value As String
    Value = ConsumeToken(Entries, ColumnIndex, "argument " & CStr(OutputIndex) & " in " & Destination.Name)

    Dim IsVariableCommand As Boolean
    IsVariableCommand = (ContextQuery.ArgumentCount = -1)

    Dim ArgumentType As QueryReturnType
    If IsVariableCommand Then
        If Not Exists(Value) Then
            Err.Raise vbObjectError + 1008, "QueryCollection", "Expected a command as argument " & CStr(OutputIndex) & " of " & Destination.Name & ": " & Value
        End If

        Dim Source As QueryDefinition
        Set Source = GetDefinition(Value)
        Set ArgumentType = Source.ReturnType
    Else
        Set ArgumentType = ContextQuery.Argument(OutputIndex).ArgumentReturn
    End If

    Dim Argument As QueryArgument
    Set Argument = QueryArgument.Create(ArgumentType)

    Dim InputIndex As Long
    If Left(Value, 1) = "$" Then
        If IsVariableCommand Then
            Err.Raise vbObjectError + 1009, "QueryCollection", "Variable command arguments must reference definitions: " & Value
        End If

        InputIndex = ParseInputIndex(Value)
        Set Argument.Link = QueryArgumentLink.Create(InputIndex, OutputIndex)
    ElseIf Exists(Value) Then
        Dim Instance As QueryDefinition
        If IsTableDefinition(Value) Then
            Call ResolveDefinition(Value)
            Set Source = GetDefinition(Value)
            Set Instance = Source.Copy()
            Call BindInstanceInputs(Instance, Destination, Entries, ColumnIndex)
        Else
            Set Instance = QueryDefinition.Create(Value, Nothing, 0)
            Call DefineCommandBody(Instance, Value, Entries, ColumnIndex)
        End If
        Set Argument.Source = Instance
    Else
        If IsVariableCommand Then
            Err.Raise vbObjectError + 1010, "QueryCollection", "Expected a command reference, not a literal: " & Value
        End If
        Argument.Value = Replace(Value, Chr(34), Empty)
    End If

    Call Destination.AddArgument(Argument)
End Sub


Private Sub BindInstanceInputs(ByVal Instance As QueryDefinition, ByVal Destination As QueryDefinition, ByVal Entries As Collection, ByRef ColumnIndex As Long)
    Dim i        As Long
    Dim Argument As QueryArgument
    Dim Value    As String
    Dim Source   As QueryDefinition
    Dim Nested   As QueryDefinition


    For i = 1 To Instance.Arguments.Count
        Set Argument = Instance.Argument(i)
        If Not Argument.Link Is Nothing Then
            Value = ConsumeToken(Entries, ColumnIndex, "exposed input " & CStr(Argument.Link.InputIndex) & " in " & Destination.Name)
            ' Replace the exposed-input link with its bound value.
            Set Argument.Link = Nothing
            Set Argument.Source = Nothing

            If Left(Value, 1) = "$" Then
                Set Argument.Link = QueryArgumentLink.Create(ParseInputIndex(Value), i)
            ElseIf Exists(Value) Then
                If IsTableDefinition(Value) Then
                    Call ResolveDefinition(Value)
                    Set Source = GetDefinition(Value)
                    Set Nested = Source.Copy()
                    Call BindInstanceInputs(Nested, Destination, Entries, ColumnIndex)
                Else
                    Set Nested = QueryDefinition.Create(Value, Nothing, 0)
                    Call DefineCommandBody(Nested, Value, Entries, ColumnIndex)
                End If
                Set Argument.Source = Nested
            Else
                Argument.Value = Replace(Value, Chr(34), Empty)
            End If

        ElseIf Not Argument.Source Is Nothing Then
            ' Continue through nested command definitions.
            Call BindInstanceInputs(Argument.Source, Destination, Entries, ColumnIndex)

        End If

    Next i
End Sub


'===============================================================================
' COMPILATION
'===============================================================================

Public Function Compile(ByVal Name As String, ParamArray Arguments() As Variant) As IQueryCommand
    Dim Args As New Collection
    Dim i As Long
    Dim ToCompile As QueryDefinition

    If Not Exists(Name) Then
        Err.Raise vbObjectError + 1100, "QueryCollection", "Unknown query definition: " & Name
    End If

    For i = LBound(Arguments) To UBound(Arguments)
        Call Args.Add(Arguments(i))
    Next i

    If IsTableDefinition(Name) Then
        Call ResolveDefinition(Name)
    End If

    Set ToCompile = GetDefinition(Name)

    Set Compile = SubCompile(ToCompile, Args)
End Function


Private Function SubCompile(ByVal ToCompile As QueryDefinition, ByVal Arguments As Collection) As IQueryCommand
    Dim ResolvedArgs As Collection
    Dim CompiledBasic As IQueryCommand
    Dim Source As QueryDefinition

    Set ResolvedArgs = ResolveArgumentLinks(ToCompile, Arguments)

    Set CompiledBasic = CompileBasic(ToCompile.BasicCommand, ResolvedArgs)

    If Not CompiledBasic Is Nothing Then
        Set SubCompile = CompiledBasic
        Exit Function
    End If

    ' Fallback for a composed definition.
    If ToCompile.ArgumentCount = 0 Then
        Err.Raise vbObjectError + 1101, "QueryCollection", "Definition cannot be compiled: " & ToCompile.Name
    End If

    Set Source = ToCompile.Argument(1).Source

    If Source Is Nothing Then
        Err.Raise vbObjectError + 1101, "QueryCollection", "Definition cannot be compiled: " & ToCompile.Name
    End If

    Set SubCompile = SubCompile(Source, Arguments)
End Function


Private Function CompileBasic(ByVal Name As String, ByVal Arguments As Collection) As IQueryCommand
    Select Case Name
        Case "Where"        : Set CompileBasic = QueryWhere.Create(Arguments.Item(1), Arguments.Item(2), Arguments.Item(3))
        Case "WhereBetween" : Set CompileBasic = QueryWhereBetween.Create(Arguments.Item(1), Arguments.Item(2), Arguments.Item(3))
        Case "WhereInList"  : Set CompileBasic = QueryWhereInList.Create(Arguments.Item(1), Arguments.Item(2))
        Case "WhereGroup"   : Set CompileBasic = QueryWhereGroup.CreateFromCollection(Arguments.Item(1), Arguments.Item(2))
        Case "OrderByField" : Set CompileBasic = QueryOrderByField.Create(Arguments.Item(1), Arguments.Item(2))
        Case "OrderBy"      : Set CompileBasic = QueryOrderBy.CreateFromCollection(Arguments)
        Case "Top"          : Set CompileBasic = QueryTop.Create(Arguments.Item(1))
        Case "Bottom"       : Set CompileBasic = QueryBottom.Create(Arguments.Item(1))
        Case "Skip"         : Set CompileBasic = QuerySkip.CreateFromCollection(Arguments)
        Case "Filter"       : Set CompileBasic = QueryFilter.Create(Arguments.Item(1))
        Case "Add"          : Set CompileBasic = QueryAdd.Create(Arguments.Item(1))
        Case "Current"      : Set CompileBasic = QueryCurrent.Create()
        Case "Source"       : Set CompileBasic = QuerySource.Create(Arguments.Item(1))
        Case "Select"       : Set CompileBasic = QuerySelect.CreateFromCollection(Arguments)
        Case "Union"        : Set CompileBasic = QueryUnion.CreateFromCollection(Arguments)
        Case "Unique"       : Set CompileBasic = QueryUnique.Create(Arguments.Item(1))
        Case "Map"          : Set CompileBasic = QueryMap.Create(Arguments.Item(1), NamedFunction.Create(Arguments.Item(2)))
        Case "Function"     : Set CompileBasic = QueryFunction.Create(Arguments.Item(1), NamedFunction.Create(Arguments.Item(2)))
        Case "Chain"        : Set CompileBasic = QueryChain.CreateFromCollection(Arguments)
        Case Else           : Set CompileBasic = Nothing
    End Select
End Function


Private Function ResolveArgumentLinks(ByVal ToCompile As QueryDefinition, ByVal Arguments As Collection) As Collection
    Dim Result As New Collection
    Dim Argument As QueryArgument
    Dim ResolvedValue As Variant
    Dim NestedCommand As IQueryCommand
    Dim i As Long

    For i = 1 To ToCompile.Arguments.Count
        Set Argument = ToCompile.Argument(i)
        If Not Argument.Link Is Nothing Then
            If Argument.Link.InputIndex < 1 Or Argument.Link.InputIndex > Arguments.Count Then
                Err.Raise vbObjectError + 1102, "QueryCollection", "Missing input argument " & CStr(Argument.Link.InputIndex) & " in definition: " & ToCompile.Name
            End If

            ResolvedValue = Arguments.Item(Argument.Link.InputIndex)
            Call Result.Add(ResolvedValue)
        ElseIf Not Argument.Source Is Nothing Then
            Set NestedCommand = SubCompile(Argument.Source, Arguments)
            Call Result.Add(NestedCommand)
        Else
            Call Result.Add(Argument.Value)
        End If
    Next i

    Set ResolveArgumentLinks = Result
End Function


'===============================================================================
' TOKENS
'===============================================================================

Private Function ParseInputIndex(ByVal nValue As String) As Long
    Dim Position As Long
    Dim IndexText As String
    Dim Character As String

    If Len(nValue) < 2 Then
        Err.Raise vbObjectError + 1006, "QueryCollection", "Invalid input reference: " & nValue
    End If

    Position = 2

    Do While Position <= Len(nValue)

        Character = Mid(nValue, Position, 1)

        If Character < "0" Or Character > "9" Then Exit Do

        IndexText = IndexText & Character
        Position = Position + 1

    Loop

    If Len(IndexText) = 0 Then
        Err.Raise vbObjectError + 1007, "QueryCollection", "Invalid input index: " & nValue
    End If

    ParseInputIndex = CLng(IndexText)
End Function


Private Function ConsumeToken(ByVal Entries As Collection, ByRef ColumnIndex As Long, ByVal Context As String) As String
    Dim Entry As FieldEntry
    Dim Value As String

    Do While ColumnIndex <= Entries.Count
        Set Entry = Entries(ColumnIndex)
        Value = Trim(CStr(Entry.DataEntry.Value))
        ColumnIndex = ColumnIndex + 1
        If Len(Value) > 0 Then
            ConsumeToken = Value
            Exit Function
        End If
    Loop
    Err.Raise vbObjectError + 1005, "QueryCollection", "Missing token (" & Context & ")"
End Function


'===============================================================================
' HELPERS
'===============================================================================

Private Function IsInCollection(ByVal nCollection As Collection, ByVal nKey As String) As Boolean
    On Error Resume Next

    Err.Clear
    Call nCollection.Item(nKey)

    IsInCollection = (Err.Number = 0)

    Err.Clear
    On Error GoTo 0
End Function

Private Sub SkipEmptyEntries(ByVal Entries As Collection)
    Dim Entry As FieldEntry
    Dim Index As Long
    Index = 1
    For Each Entry In Entries
        If Trim(CStr(Entry.DataEntry.Value)) = Empty Then
            Call Entries.Remove(Index)
        Else
            Index = Index + 1
        End If
    Next Entry
End Sub


'===============================================================================
' BASIC DEFINITIONS
'===============================================================================

Private Sub PreDefineBasics()
    Dim tScalar     As QueryReturnType
    Dim tSequence   As QueryReturnType
    Dim tDataSet    As QueryReturnType
    Dim tFunction   As QueryReturnType
    Dim tCollection As QueryReturnType

    Set tScalar     = QueryReturnType.Create("Scalar", QueryValueKindScalar)
    Set tSequence   = QueryReturnType.Create("Sequence", QueryValueKindSequence)
    Set tDataSet    = QueryReturnType.Create("DataSet", QueryValueKindDataSet)
    Set tFunction   = QueryReturnType.Create("Function", QueryValueKindFunction)
    Set tCollection = QueryReturnType.Create("Collection", QueryValueKindSequence)

    Call AddBasicDefinition("Where"        , tSequence  , Array(tScalar, tScalar, tScalar) ,     3)
    Call AddBasicDefinition("WhereBetween" , tSequence  , Array(tScalar, tScalar, tScalar) ,     3)
    Call AddBasicDefinition("WhereInList"  , tSequence  , Array(tScalar, tCollection)      ,     2)
    Call AddBasicDefinition("WhereGroup"   , tSequence  , Array(tScalar, tCollection)      ,     -1)
    Call AddBasicDefinition("Filter"       , tDataSet   , Array(tSequence)                 ,     1)
    Call AddBasicDefinition("OrderByField" , tDataSet   , Array(tScalar, tScalar)          ,     2)
    Call AddBasicDefinition("OrderBy"      , tDataSet   , Array(tCollection)               ,     -1)
    Call AddBasicDefinition("Add"          , tDataSet   , Array(tScalar)                   ,     1)
    Call AddBasicDefinition("Bottom"       , tDataSet   , Array(tScalar)                   ,     1)
    Call AddBasicDefinition("Top"          , tDataSet   , Array(tScalar)                   ,     1)
    Call AddBasicDefinition("Select"       , tDataSet   , Array(tCollection)               ,     -1)
    Call AddBasicDefinition("Skip"         , tDataSet   , Array(tScalar)                   ,     1)
    Call AddBasicDefinition("Union"        , tDataSet   , Array(tCollection)               ,     -1)
    Call AddBasicDefinition("Unique"       , tDataSet   , Array(tScalar)                   ,     1)
    Call AddBasicDefinition("Source"       , tDataSet   , Array(tScalar)                   ,     1)
    Call AddBasicDefinition("Current"      , tDataSet   , Array()                          ,     0)
    Call AddBasicDefinition("Map"          , tSequence  , Array(tScalar, tScalar)          ,     2)
    Call AddBasicDefinition("Function"     , tScalar    , Array(tScalar, tScalar)          ,     2)
    Call AddBasicDefinition("Chain"        , tScalar    , Array(tCollection)               ,     -1)

    Call mResolved.Add(True, "Where")
    Call mResolved.Add(True, "WhereBetween")
    Call mResolved.Add(True, "WhereInList")
    Call mResolved.Add(True, "WhereGroup")
    Call mResolved.Add(True, "Filter")
    Call mResolved.Add(True, "Order")
    Call mResolved.Add(True, "OrderBy")
    Call mResolved.Add(True, "Add")
    Call mResolved.Add(True, "Bottom")
    Call mResolved.Add(True, "Top")
    Call mResolved.Add(True, "Select")
    Call mResolved.Add(True, "Skip")
    Call mResolved.Add(True, "Union")
    Call mResolved.Add(True, "Unique")
    Call mResolved.Add(True, "Source")
    Call mResolved.Add(True, "Current")
    Call mResolved.Add(True, "Map")
    Call mResolved.Add(True, "Function")
    Call mResolved.Add(True, "Chain")
End Sub


Private Sub AddBasicDefinition(ByVal nName As String, ByVal nReturnType As QueryReturnType, ByVal nArguments As Variant, ByVal nArgumentCount As Long)
    Dim Definition As QueryDefinition
    Dim ArgumentType As QueryReturnType
    Dim Argument As QueryArgument
    Dim i As Long
    Dim LowerIndex As Long

    Set Definition = QueryDefinition.Create(nName, nReturnType, nArgumentCount)
    Call Definition.SetBasicCommand(nName)

    If nArgumentCount > 0 Then

        LowerIndex = LBound(nArguments)

        For i = 1 To nArgumentCount
            Set ArgumentType = nArguments(LowerIndex + i - 1)
            Set Argument = QueryArgument.Create(ArgumentType)
            Call Definition.AddArgument(Argument)
        Next i

    End If

    Call AddBasicDefinitionToCollection(Definition)
End Sub


Private Sub AddBasicDefinitionToCollection(ByVal Definition As QueryDefinition)
    Call mDefinitions.Add(Definition, Definition.Name)
End Sub
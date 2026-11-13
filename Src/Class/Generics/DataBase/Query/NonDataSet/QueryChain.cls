VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IQueryCommand

Private mCommands As Collection


'===============================================================================
' CREATION
'===============================================================================

Private Sub Class_Initialize()
    Set mCommands = New Collection
End Sub


'===============================================================================
' COMMANDS
'===============================================================================

Public Function CreateFromCollection(ByVal Commands As Collection) As QueryChain
    Dim Result As QueryChain
    Dim Command As IQueryCommand

    Set Result = New QueryChain
    For Each Command In Commands
        Call Result.Add(Command)
    Next Command
    Set CreateFromCollection = Result
End Function

Public Function Create(ParamArray Commands() As Variant) As QueryChain
    Dim Result As QueryChain
    Dim i As Long

    Set Result = New QueryChain
    For i = LBound(Commands) To UBound(Commands)
        Call Result.Add(Commands(i))
    Next i
    Set Create = Result
End Function

Public Sub Add(ByVal Command As IQueryCommand)
    If Command Is Nothing Then
        Err.Raise vbObjectError + 1300, _
                  "QueryChain.Add", _
                  "QueryCommand cannot be Nothing."
    End If
    Call mCommands.Add(Command)
End Sub


Public Property Get Commands() As Collection
    Set Commands = mCommands
End Property


Public Property Get Count() As Long
    Count = mCommands.Count
End Property


Public Property Get Command(ByVal Index As Long) As IQueryCommand
    Set Command = mCommands(Index)
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    Dim Command As IQueryCommand
    Dim Result  As QueryValue

    For Each Command In mCommands
        Set Result = Command.Execute(Context)
        If Result Is Nothing Then
            Err.Raise vbObjectError + 1301, _
                      "QueryChain.Execute", _
                      "QueryCommand returned no DataSet."
        End If
        Set Context.Current = Result
    Next Command

    Set IQueryCommand_Execute = Context.Current
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryCommandChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IQueryCommand

Private mCommands As Collection


'===============================================================================
' INITIALIZATION
'===============================================================================

Private Sub Class_Initialize()
    Set mCommands = New Collection
End Sub


'===============================================================================
' COMMANDS
'===============================================================================

Public Function Create(ParamArray Commands() As Variant) As QueryCommandChain
    Set Create = New QueryCommandChain
    Dim i As Long
    For i = 0 To Ubound(Commands)
        Call Create.Add(Commands(i))
    Next i
End Function

Public Sub Add(ByVal Command As IQueryCommand)
    If Command Is Nothing Then
        Err.Raise vbObjectError + 1300, _
                  "QueryCommandChain.Add", _
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

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Command As IQueryCommand
    Dim Result As DataSet

    For Each Command In mCommands
        Set Result = Command.Execute(Context)
        If Result Is Nothing Then
            Err.Raise vbObjectError + 1301, _
                      "QueryCommandChain.Execute", _
                      "QueryCommand returned no DataSet."
        End If
        Set Context.Current = Result
    Next Command

    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1302, _
                  "QueryCommandChain.Execute", _
                  "QueryCommandChain produced no DataSet."
    End If
    Set IQueryCommand_Execute = Context.Current
End Function
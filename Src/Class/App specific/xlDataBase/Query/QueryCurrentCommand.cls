VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryCurrentCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1400, _
                  "QueryCurrentCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set IQueryCommand_Execute = Context.Current
End Function
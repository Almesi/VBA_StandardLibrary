VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryBottomCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mCount As Long


'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal Count As Long) As QueryBottomCommand
    Set Create = New QueryBottomCommand
    Call Create.Initialize(Count)
End Function


Friend Sub Initialize(ByVal Count As Long)
    If Count < 0 Then
        Err.Raise vbObjectError + 1450, _
                  "QueryBottomCommand.Initialize", _
                  "Count cannot be negative."
    End If

    Let mCount = Count
End Sub


'===============================================================================
' PROPERTIES
'===============================================================================

Public Property Get Count() As Long
    Count = mCount
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source As DataSet
    Dim Result As DataSet
    Dim RowIndex As Long
    Dim StartIndex As Long

    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1451, _
                  "QueryBottomCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set Source = Context.Current
    Set Result = New DataSet

    Call CopyStructure(Source, Result)

    If mCount = 0 Then
        Set IQueryCommand_Execute = Result
        Exit Function
    End If

    StartIndex = Source.RowCount - mCount + 1

    If StartIndex < 1 Then
        StartIndex = 1
    End If

    For RowIndex = StartIndex To Source.RowCount
        Call Result.AddRow(Source.Row(RowIndex))
    Next RowIndex

    Set IQueryCommand_Execute = Result
End Function


'===============================================================================
' COPY STRUCTURE
'===============================================================================

Private Sub CopyStructure(ByVal Source As DataSet, ByVal Result As DataSet)
    Dim Index As Long

    For Index = 1 To Source.FieldCount
        Call Result.AddFieldDefinition(Source.FieldDefinition(Index))
    Next Index
End Sub
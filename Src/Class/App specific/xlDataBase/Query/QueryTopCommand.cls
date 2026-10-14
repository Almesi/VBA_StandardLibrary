VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryTopCommand"
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

Public Function Create(ByVal Count As Long) As QueryTopCommand
    Set Create = New QueryTopCommand
    Call Create.Initialize(Count)
End Function


Friend Sub Initialize(ByVal Count As Long)
    If Count < 0 Then
        Err.Raise vbObjectError + 1440, _
                  "QueryTopCommand.Initialize", _
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
    Dim Limit As Long

    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1441, _
                  "QueryTopCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set Source = Context.Current
    Set Result = New DataSet

    Call CopyStructure(Source, Result)

    Limit = mCount

    If Limit > Source.RowCount Then
        Limit = Source.RowCount
    End If

    For RowIndex = 1 To Limit
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
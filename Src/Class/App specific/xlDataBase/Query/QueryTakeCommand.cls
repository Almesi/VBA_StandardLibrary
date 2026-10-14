VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryTakeCommand"
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

Public Function Create(ByVal Count As Long) As QueryTakeCommand
    Set Create = New QueryTakeCommand
    Call Create.Initialize(Count)
End Function

Friend Sub Initialize(ByVal Count As Long)
    If Count < 0 Then
        Err.Raise vbObjectError + 1420, _
                  "QueryTakeCommand.Initialize", _
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

    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1421, _
                  "QueryTakeCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set Source = Context.Current
    Set Result = New DataSet
    Call CopyStructure(Source, Result)
    For RowIndex = 1 To MinLong(mCount, Source.RowCount)
        Call CopyRow(Source, Result, RowIndex)
    Next RowIndex
    Set IQueryCommand_Execute = Result
End Function


'===============================================================================
' HELPERS
'===============================================================================

Private Function MinLong(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    If Value1 < Value2 Then
        MinLong = Value1
    Else
        MinLong = Value2
    End If
End Function


Private Sub CopyStructure(ByVal Source As DataSet, ByVal Result As DataSet)
    Dim Index As Long

    For Index = 1 To Source.FieldDefinitionCount
        Call Result.AddFieldDefinition(Source.FieldDefinition(Index))
    Next Index
End Sub


Private Sub CopyRow(ByVal Source As DataSet, ByVal Result As DataSet, ByVal RowIndex As Long)
    Dim FieldIndex As Long
    Dim DataEntry As DataEntry
    For FieldIndex = 1 To Source.FieldCount
        Set DataEntry = Source.DataEntryAt(RowIndex, FieldIndex)
        Call Result.AddDataEntry(DataEntry)
    Next FieldIndex
End Sub
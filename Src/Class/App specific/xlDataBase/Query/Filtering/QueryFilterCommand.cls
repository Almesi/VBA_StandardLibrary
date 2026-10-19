VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryFilterCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IQueryCommand

Private mFilter  As IQueryCommand


'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal Filter As IQueryCommand) As QueryFilterCommand
    Set Create = New QueryFilterCommand
    Call Create.Initialize(Filter)
End Function

Friend Sub Initialize(ByVal Filter As IQueryCommand)
    Set mFilter  = Filter
End Sub


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    If Not Context.HasCurrent Then Err.Raise vbObjectError + 1453, "QueryWhereInListCommand.Execute", "QueryContext has no current DataSet."

    Dim Source  As DataSet
    Dim Filters As DataSet

    Set Source = Context.Current
    Set Filters = mFilter.Execute(Context)
    Set IQueryCommand_Execute = FilterDataSet(Filters, Source)
End Function

Private Function FilterDataSet(ByVal Filters As DataSet, ByVal Source As DataSet) As DataSet
    If Filters.RowCount <> Source.RowCount Then Err.Raise vbObjectError + 1465, "QueryFilterCommand.FilterDataSet", "Filter row count must match source row count."

    Dim MatchField As Field
    Set MatchField = Filters.FieldByName("Match")

    If MatchField Is Nothing Then Err.Raise vbObjectError + 1466, "QueryFilterCommand.FilterDataSet", "Filter dataset must contain a 'Match' field."

    Dim Result As DataSet
    Set Result = Source.CopyStructure

    Dim RowIndex As Long
    For RowIndex = 1 To Source.RowCount
        Dim MatchEntry As DataEntry
        Set MatchEntry = MatchField.Item(RowIndex)

        If MatchEntry.Value Then
            Call Result.AddRow(Source.Row(RowIndex))
        End If
    Next RowIndex

    Set FilterDataSet = Result
End Function
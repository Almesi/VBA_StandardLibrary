VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IQueryCommand

Private mFilter  As IQueryCommand


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal Filter As IQueryCommand) As QueryFilter
    Set Create = New QueryFilter
    Call Create.Initialize(Filter)
End Function

Friend Sub Initialize(ByVal Filter As IQueryCommand)
    Set mFilter  = Filter
End Sub


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet) = False Then Exit Function

    Dim Source  As DataSet
    Set Source = Context.Current.Value

    Dim Predicate As QueryValue
    Set Predicate = mFilter.Execute(Context)
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, FilterDataSet(Predicate, Source))
End Function

Private Function FilterDataSet(ByVal Predicate As QueryValue, ByVal Source As DataSet) As DataSet
    If Predicate.Kind        <> QueryValueKindSequence Then Err.Raise vbObjectError + 1465, "QueryFilter.FilterDataSet", "Not correct QueryValue."
    If Predicate.Value.Count <> Source.RowCount        Then Err.Raise vbObjectError + 1465, "QueryFilter.FilterDataSet", "Filter row count must match source row count."

    Dim Matches As Variant
    Set Matches = Predicate.Value

    Dim Result As DataSet
    Set Result = Source.CopyStructure

    Dim RowIndex As Long
    For RowIndex = 1 To Source.RowCount
        If Matches.Item(RowIndex).Value Then
            Call Result.AddRow(Source.Row(RowIndex))
        End If
    Next RowIndex

    Set FilterDataSet = Result
End Function
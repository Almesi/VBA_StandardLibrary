VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryWhereInList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mFieldName As String
Private mValues    As Collection

'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal FieldName As String, ByVal Values As Collection) As QueryWhereInList
    Set Create = New QueryWhereInList
    Call Create.Initialize(FieldName, Values)
End Function

Friend Sub Initialize(ByVal FieldName As String, ByVal Values As Collection)
    Let mFieldName = FieldName
    Set mValues    = Values
End Sub

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet, mFieldName) = False Then Exit Function

    Dim Field As Field
    Set Field = Context.Current.Value.FieldByName(mFieldName)

    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindSequence, EvaluateField(Field))
End Function

Private Function EvaluateField(ByVal Field As Field)As Collection
    Dim Result As New Collection
    Dim Entry As DataEntry

    For Each Entry In Field.DataEntries
        Call Result.Add(DataEntry.Create(Matches(Entry)))
    Next Entry

    Set EvaluateField = Result
End Function

Private Function Matches(ByVal Entry As DataEntry) As Boolean
    Dim Item As Variant

    For Each Item In mValues
        If Entry.EqualsValue(Item) Then
            Matches = True
            Exit Function
        End If
    Next Item
End Function
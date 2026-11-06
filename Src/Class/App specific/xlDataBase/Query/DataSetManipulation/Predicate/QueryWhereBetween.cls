VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryWhereBetween"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mFieldName  As String
Private mLowerValue As Variant
Private mUpperValue As Variant

'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal FieldName As String, ByVal LowerValue As Variant, ByVal UpperValue As Variant) As QueryWhereBetween
    Set Create = New QueryWhereBetween
    Call Create.Initialize(FieldName, LowerValue, UpperValue)
End Function

Friend Sub Initialize(ByVal FieldName As String, ByVal LowerValue As Variant, ByVal UpperValue As Variant)
    mFieldName  = FieldName
    mLowerValue = LowerValue
    mUpperValue = UpperValue
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
    Matches = (Entry.IsGreaterThanOrEqual(mLowerValue) And Entry.IsLessThanOrEqual(mUpperValue))
End Function
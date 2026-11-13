VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mFieldName  As String
Private mFunc       As IFunction


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal nFieldName As String, ByVal nFunc As IFunction) As QueryMap
    Set Create = New QueryMap
    Call Create.Initialize(nFieldName, nFunc)
End Function

Friend Sub Initialize(ByVal nFieldName As String, ByVal nFunc As IFunction)
    Let mFieldName = FieldName
    Set mFunc      = nFunc
End Sub


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet, mFieldName) = False Then Exit Function

    Dim Source As DataSet
    Set Source = Context.Current.Value
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindSequence, MapDataSet(Source))
End Function

Private Function MapDataSet(ByVal Source As DataSet) As Collection
    Dim Result As New Collection

    Dim Field As Field
    Set Field = Result.FieldByName(mFieldName)

    Dim Entries As Collection
    Set Entries = Field.DataEntries

    Dim Entry As DataEntry
    For Each Entry In Entries
        Call Result.Add(mFunc.Evaluate(Entry.Value))
    Next Entry

    Set MapDataSet = Result
End Function
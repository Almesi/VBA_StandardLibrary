VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryOrderByField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mFieldName As String
Private mDescending As Boolean


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal FieldName As String, Optional ByVal Descending As Boolean = False) As QueryOrderByField
    Set Create = New QueryOrderByField
    Call Create.Initialize(FieldName, Descending)
End Function


Friend Sub Initialize(ByVal FieldName As String, ByVal Descending As Boolean)
    Let mFieldName  = FieldName
    Let mDescending = Descending
End Sub

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet, mFieldName) = False Then Exit Function

    Dim Source As DataSet
    Set Source = Context.Current.Value
    Call ValidateFields(Source)

    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, SortDataSet(Source))
End Function

'TODO Implement
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mCount As Long


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal Count As Long) As QueryTop
    Set Create = New QueryTop
    Call Create.Initialize(Count)
End Function


Friend Sub Initialize(ByVal Count As Long)
    Let mCount = Count
End Sub

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet) = False Then Exit Function

    Dim Source As DataSet
    Set Source = Context.Current.Value

    Dim Result As DataSet
    Set Result = Source.CopyStructure

    Dim Limit As Long
    Limit = mCount
    If Limit > Source.RowCount Then
        Limit = Source.RowCount
    End If

    Dim RowIndex As Long
    For RowIndex = 1 To Limit
        Call Result.AddRow(Source.Row(RowIndex))
    Next RowIndex

    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, Result)
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mAddName As String


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal AddName As String) As QueryAdd
    Set Create = New QueryAdd
    Call Create.Initialize(AddName)
End Function

Friend Sub Initialize(ByVal AddName As String)
    If Len(Trim(AddName)) = 0 Then

        Err.Raise vbObjectError + 1401, _
                  "QueryAdd.Initialize", _
                  "Add name cannot be empty."
    End If
    Let mAddName = AddName
End Sub

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet, mAddName) = False Then Exit Function

    Dim Result As DataSet
    Set Result = Context.Current.Value
    Call Context.Add(NamedDataSet.Create(AddName, Result))
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, Result)
End Function
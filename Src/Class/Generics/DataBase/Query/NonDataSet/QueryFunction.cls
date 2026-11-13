VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryFunction"
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

Public Function Create(ByVal nFieldName As String, ByVal nFunc As IFunction) As QueryFunction
    Set Create = New QueryFunction
    Call Create.Initialize(nFieldName, nFunc)
End Function

Friend Sub Initialize(ByVal nFieldName As String, ByVal nFunc As IFunction)
    Let mFieldName = nFieldName
    Set mFunc      = nFunc
End Sub


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet, mFieldName) = False Then Exit Function

    Dim Source As DataSet
    Set Source = Context.Current.Value
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindScalar, FunctionDataSet(Source))
End Function

Private Function FunctionDataSet(ByVal Source As DataSet) As Variant
    Dim Field As Field
    Set Field = Source.FieldByName(mFieldName)

    Dim Entries As Collection
    Set Entries = Field.DataEntries

    Dim Size As Long
    Size = Entries.Count - 1
        
    Dim Arr() As Variant
    ReDim Arr(Size)
    
    Dim i As Long
    For i = 0 To Size
        Call Assign(Arr(i), Entries.Item(i + 1).Value)
    Next i

    Call Assign(FunctionDataSet, mFunc.Evaluate(Arr))
End Function
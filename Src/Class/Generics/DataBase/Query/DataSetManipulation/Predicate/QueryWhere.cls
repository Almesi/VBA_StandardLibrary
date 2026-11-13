VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IQueryCommand

Private mFieldName As String
Private mOperator  As String
Private mValue     As Variant


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal FieldName As String, ByVal Operator As String, ByVal Value As Variant) As QueryWhere
    Set Create = New QueryWhere
    Call Create.Initialize(FieldName, Operator, Value)
End Function

Friend Sub Initialize(ByVal FieldName As String, ByVal Operator As String, ByVal Value As Variant)
    Let mFieldName = FieldName
    Let mOperator  = UCase(Trim(Operator))
    Let mValue     = Value
End Sub

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet, mFieldName) = False Then Exit Function
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindSequence, FilterDataSet(Context.Current.Value))
End Function

Private Function FilterDataSet(ByVal Source As DataSet)As Collection
    Dim Field As Field
    Set Field = Source.FieldByName(mFieldName)

    Dim Result As New Collection
    Dim Entry      As DataEntry
    For Each Entry In Field.DataEntries
        Call Result.Add(DataEntry.Create(Compare(Entry)))
    Next Entry

    Set FilterDataSet = Result
End Function

Private Function Compare(ByVal LeftEntry As DataEntry) As Boolean
    Select Case mOperator
        Case "="           : Compare = LeftEntry.EqualsValue(mValue)
        Case "<>"          : Compare = Not LeftEntry.EqualsValue(mValue)
        Case ">"           : Compare = LeftEntry.IsGreaterThan(mValue)
        Case ">="          : Compare = LeftEntry.IsGreaterThanOrEqual(mValue)
        Case "<"           : Compare = LeftEntry.IsLessThan(mValue)
        Case "<="          : Compare = LeftEntry.IsLessThanOrEqual(mValue)
        Case "LIKE"        : Compare = CStr(LeftEntry.Value) Like CStr(mValue)
        Case "IS NULL"     : Compare = LeftEntry.Equals(mValue)
        Case "IS NOT NULL" : Compare = Not LeftEntry.Equals(mValue)
        Case Else
            Err.Raise vbObjectError + 1414, _
                      "QueryWhereCommand.Compare", _
                      "Unsupported operator '" & mOperator & "'."
    End Select
End Function
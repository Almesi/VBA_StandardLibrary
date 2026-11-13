VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QuerySelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mFields As Collection


'===============================================================================
' CREATION
'===============================================================================

Private Sub Class_Initialize()
    Set mFields = New Collection
End Sub

Public Function CreateFromCollection(ByVal Names As Collection) As QuerySelect
    Dim Result As QuerySelect
    Dim Name As String

    Set Result = New QuerySelect
    For Each Name In Names
        Call Result.Add(Name)
    Next Name
    Set CreateFromCollection = Result
End Function

Public Function Create(ParamArray Names() As Variant) As QuerySelect
    Dim Result As QuerySelect
    Dim i As Long

    Set Result = New QuerySelect
    For i = LBound(Names) To UBound(Names)
        Call Result.Add(Names(i))
    Next i
    Set Create = Result
End Function



Public Sub Add(ByVal Name As String)
    If Len(Trim(Name)) = 0 Then Exit Sub
    Call mFields.Add(Name)
End Sub


Public Property Get Fields() As Collection
    Set Fields = mFields
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet) = False Then Exit Function

    Dim Source As DataSet
    Set Source = Context.Current.Value

    Dim Result As DataSet
    Set Result = DataSet.Create

    Dim FieldName   As Variant
    For Each FieldName In mFields
        If QueryAssert(Context, QueryValueKindDataSet, FieldName) = False Then Exit Function
        Dim SourceField As Field
        Set SourceField = Source.FieldByName(CStr(FieldName))
        Call Result.AddField(SourceField)
    Next FieldName
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, Result)
End Function
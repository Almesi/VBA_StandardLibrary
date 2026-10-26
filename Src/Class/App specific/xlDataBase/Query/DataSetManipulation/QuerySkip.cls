VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QuerySkip"
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

Public Function Create(ParamArray Names() As Variant) As QuerySkip
    Set Create = New QuerySkip
    Dim i As Long
    For i = 0 To Ubound(Names)
        Call Create.Add(Names(i))
    Next i
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

    Dim FieldItem   As Field
    For Each FieldItem In Source.Fields
        Dim SourceField As Field
        Set SourceField = SkipFieldName(FieldItem)

        If Not SourceField Is Nothing Then
            Call Result.AddField(SourceField)
        End If
    Next FieldItem

    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, Result)
End Function

Private Function SkipFieldName(ByVal SourceField As Field) As Field
    Dim FieldName As Variant

    For Each FieldName In mFields
        If StrComp(CStr(FieldName), SourceField.Definition.Name, vbTextCompare) = 0 Then
            Exit Function
        End If
    Next FieldName

    Set SourceFieldIfNotSkipped = SourceField
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mDataSets As Collection


'===============================================================================
' CREATION
'===============================================================================

Private Sub Class_Initialize()
    Set mDataSets = New Collection
End Sub

Public Function CreateFromCollection(ByVal Names As Collection) As QueryUnion
    Dim Result As QueryUnion
    Dim Name As String

    Set Result = New QueryUnion
    For Each Name In Names
        Call Result.Add(Name)
    Next Name
    Set CreateFromCollection = Result
End Function

Public Function Create(ParamArray Names() As Variant) As QueryUnion
    Dim Result As QueryUnion
    Dim i As Long

    Set Result = New QueryUnion
    For i = LBound(Names) To UBound(Names)
        Call Result.Add(Names(i))
    Next i
    Set Create = Result
End Function


Public Sub Add(ByVal Name As String)
    If Len(Trim(Name)) = 0 Then Exit Sub
    Call mDataSets.Add(Name)
End Sub


Public Property Get DataSets() As Collection
    Set DataSets = mDataSets
End Property


'===============================================================================
' EXECUTE
'===============================================================================
Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet) = False Then Exit Function

    Dim Result As DataSet
    Set Result = Context.Current.Value

    Dim DataSetName As Variant
    For Each DataSetName In mDataSets
        Dim Source As DataSet
        Set Source = Context.DataSetByName(CStr(DataSetName))

        If Source Is Nothing Then
            Err.Raise vbObjectError + 1404, _
                      "QueryUnion.Execute", _
                      "DataSet '" & CStr(DataSetName) & "' was not found."
        End If

        Set Result = DataSet.Merge(Result, Source)
    Next DataSetName
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, Result)
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryUnique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand


'===============================================================================
' CREATION
'===============================================================================

Public Function Create() As QueryUnique
    Set Create = New QueryUnique
End Function


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet) = False Then Exit Function

    Dim Source As DataSet
    Set Source = Context.Current.Value
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, UniqueDataSet(Source))
End Function


Private Function UniqueDataSet(ByVal Source As DataSet) As DataSet
    Dim Result   As DataSet
    Dim RowIndex As Long
    Dim RowKey   As String
    Dim Keys     As Object

    Set Result = DataSet.Create
    Set Keys = CreateObject("Scripting.Dictionary")

    Set Result = Source.CopyStructure

    For RowIndex = 1 To Source.RowCount
        RowKey = BuildRowKey(Source, RowIndex)

        If Not Keys.Exists(RowKey) Then
            Call Keys.Add(RowKey, True)
            Call Result.AddRow(Source.Row(RowIndex))
        End If
    Next RowIndex

    Set UniqueDataSet = Result
End Function

Private Function BuildRowKey(ByVal Source As DataSet, ByVal RowIndex As Long) As String
    Dim FieldIndex As Long
    Dim Value      As Variant
    Dim Key        As String

    For FieldIndex = 1 To Source.FieldCount
        Value = Source.DataEntryAt(RowIndex, FieldIndex).Value
        Key = Key & "|" & ValueKey(Value)
    Next FieldIndex

    BuildRowKey = Key
End Function

Private Function ValueKey(ByVal Value As Variant) As String
    If IsNull(Value) Then
        ValueKey = "<NULL>"
    ElseIf IsEmpty(Value) Then
        ValueKey = "<EMPTY>"
    Else
        ValueKey = TypeName(Value) & ":" & CStr(Value)
    End If
End Function
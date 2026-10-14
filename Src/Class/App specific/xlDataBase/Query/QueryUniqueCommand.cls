VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryUniqueCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand


'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create() As QueryUniqueCommand
    Set Create = New QueryUniqueCommand
End Function


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source As DataSet

    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1420, _
                  "QueryUniqueCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set Source = Context.Current
    Set IQueryCommand_Execute = UniqueDataSet(Source)
End Function


'===============================================================================
' UNIQUE
'===============================================================================

Private Function UniqueDataSet(ByVal Source As DataSet) As DataSet
    Dim Result As DataSet
    Dim RowIndex As Long
    Dim RowKey As String
    Dim Keys As Object

    Set Result = New DataSet
    Set Keys = CreateObject("Scripting.Dictionary")

    Call CopyStructure(Source, Result)

    For RowIndex = 1 To Source.RowCount
        RowKey = BuildRowKey(Source, RowIndex)

        If Not Keys.Exists(RowKey) Then
            Call Keys.Add(RowKey, True)
            Call Result.AddRow(Source.Row(RowIndex))
        End If
    Next RowIndex

    Set UniqueDataSet = Result
End Function


'===============================================================================
' BUILD ROW KEY
'===============================================================================

Private Function BuildRowKey(ByVal Source As DataSet, ByVal RowIndex As Long) As String
    Dim FieldIndex As Long
    Dim Value As Variant
    Dim Key As String

    For FieldIndex = 1 To Source.FieldCount
        Value = Source.DataEntryAt(RowIndex, FieldIndex).Value
        Key = Key & "|" & ValueKey(Value)
    Next FieldIndex

    BuildRowKey = Key
End Function


'===============================================================================
' VALUE KEY
'===============================================================================

Private Function ValueKey(ByVal Value As Variant) As String
    If IsNull(Value) Then
        ValueKey = "<NULL>"
    ElseIf IsEmpty(Value) Then
        ValueKey = "<EMPTY>"
    Else
        ValueKey = TypeName(Value) & ":" & CStr(Value)
    End If
End Function


'===============================================================================
' COPY STRUCTURE
'===============================================================================

Private Sub CopyStructure(ByVal Source As DataSet, ByVal Result As DataSet)
    Dim Index As Long

    For Index = 1 To Source.FieldCount
        Call Result.AddFieldDefinition(Source.FieldDefinition(Index))
    Next Index
End Sub
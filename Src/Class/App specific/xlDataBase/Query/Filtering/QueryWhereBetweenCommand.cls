VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryWhereBetweenCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mFieldName  As String
Private mLowerValue As Variant
Private mUpperValue As Variant

'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal FieldName As String, ByVal LowerValue As Variant, ByVal UpperValue As Variant) As QueryWhereBetweenCommand
    Set Create = New QueryWhereBetweenCommand
    Call Create.Initialize(FieldName, LowerValue, UpperValue)
End Function

Friend Sub Initialize(ByVal FieldName As String, ByVal LowerValue As Variant, ByVal UpperValue As Variant)
    mFieldName  = FieldName
    mLowerValue = LowerValue
    mUpperValue = UpperValue
End Sub

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source As DataSet
    Dim Field As Field

    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1443, "QueryWhereBetweenCommand.Execute", "QueryContext has no current DataSet."
    End If

    Set Source = Context.Current
    Set Field = Source.FieldByName(mFieldName)

    If Field Is Nothing Then
        Err.Raise vbObjectError + 1444, "QueryWhereBetweenCommand.Execute", "Field '" & mFieldName & "' was not found."
    End If

    Set IQueryCommand_Execute = EvaluateField(Field)
End Function

'===============================================================================
' EVALUATE FIELD
'===============================================================================

Private Function EvaluateField(ByVal Field As Field) As DataSet
    Dim Entries As New Collection
    Dim Entry As DataEntry

    For Each Entry In Field.DataEntries
        Call Entries.Add(DataEntry.Create(Matches(Entry)))
    Next Entry

    Dim Definition As FieldDefinition
    Set Definition = FieldDefinition.Create("Match", vbBoolean, Validator.CreateFromCoreValueType(CoreBoolean))

    Dim ResultField As Field
    Set ResultField = Field.Create(Definition, Entries)

    Dim Result As DataSet
    Set Result = DataSet.Create()
    Call Result.AddField(ResultField)

    Set EvaluateField = Result
End Function

'===============================================================================
' MATCH
'===============================================================================

Private Function Matches(ByVal Entry As Entry) As Boolean
    Matches = (Entry.IsGreaterThanOrEqual(mLowerValue) And Entry.IsLessThanOrEqual(mUpperValue))
End Function
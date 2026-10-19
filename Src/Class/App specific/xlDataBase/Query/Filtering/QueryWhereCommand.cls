VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryWhereCommand"
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
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal FieldName As String, ByVal Operator As String, ByVal Value As Variant) As QueryWhereCommand
    Set Create = New QueryWhereCommand
    Call Create.Initialize(FieldName, Operator, Value)
End Function

Friend Sub Initialize(ByVal FieldName As String, ByVal Operator As String, ByVal Value As Variant)
    If Len(Trim(FieldName)) = 0 Then
        Err.Raise vbObjectError + 1410, _
                  "QueryWhereCommand.Initialize", _
                  "Field name cannot be empty."
    End If
    If Len(Trim(Operator)) = 0 Then
        Err.Raise vbObjectError + 1411, _
                  "QueryWhereCommand.Initialize", _
                  "Operator cannot be empty."
    End If

    Let mFieldName = FieldName
    Let mOperator = UCase(Trim(Operator))
    Let mValue = Value
End Sub

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source As DataSet
    Dim Field As Field

    If Not Context.HasCurrent Then Err.Raise vbObjectError + 1453, "QueryWhereInListCommand.Execute", "QueryContext has no current DataSet."

    Set Source = Context.Current
    Set Field = Source.FieldByName(mFieldName)

    If Field Is Nothing Then Err.Raise vbObjectError + 1454, "QueryWhereInListCommand.Execute", "Field '" & mFieldName & "' was not found."

    Set IQueryCommand_Execute = FilterDataSet(Source)
End Function

Private Function FilterDataSet(ByVal Source As DataSet) As DataSet
    Dim Field As Field
    Set Field = Source.FieldByName(mFieldName)

    Dim Entries As New Collection
    Dim Entry      As DataEntry
    For Each Entry In Field.DataEntries
        Call Entries.Add(DataEntry.Create(Compare(Entry)))
    Next Entry


    ' Result Collection
    Dim Definition As FieldDefinition
    Set Definition = FieldDefinition.Create("Match", vbBoolean, Validator.CreateFromCoreValueType(CoreBoolean))

    Dim ResultField As Field
    Set ResultField = Field.Create(Definition, Entries)

    Dim Result     As DataSet
    Set Result = DataSet.Create()
    Call Result.AddField(ResultField)

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
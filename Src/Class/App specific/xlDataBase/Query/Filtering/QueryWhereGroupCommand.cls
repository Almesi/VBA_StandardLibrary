VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryWhereGroupCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Public Enum WhereGroupOperator
    Andd  = 0
    Orr   = 1
    NAndd = 2
    NOrr  = 3
    XOrr  = 4
    XNorr = 5
End Enum

Private mConditions      As Collection
Private mLogicalOperator As WhereGroupOperator

'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal Operator As WhereGroupOperator) As QueryWhereGroupCommand
    Set Create = New QueryWhereGroupCommand
    Call Create.Initialize(Operator)
End Function

Friend Sub Initialize(ByVal Operator As WhereGroupOperator)
    Let mLogicalOperator = Operator
    Set mConditions      = New Collection
End Sub

Public Property Get WhereGroupOperator() As WhereGroupOperator
    WhereGroupOperator = mLogicalOperator
End Property

'===============================================================================
' ADD CONDITION
'===============================================================================

Public Function AddCondition(ByVal Command As IQueryCommand) As QueryWhereGroupCommand
    Call mConditions.Add(Command)
    Set AddCondition = Me
End Function

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    If Not Context.HasCurrent Then Err.Raise vbObjectError + 1453, "QueryWhereGroupCommand.Execute", "QueryContext has no current DataSet."
    If mConditions Is Nothing Then Err.Raise vbObjectError + 1461, "QueryWhereGroupCommand.Execute", "Conditions have not been initialized."
    If mConditions.Count = 0  Then Err.Raise vbObjectError + 1461, "QueryWhereGroupCommand.Execute", "At least one condition is required."

    Dim Source As DataSet
    Set Source = Context.Current

    ' Evaluate the first condition to seed the accumulated match.
    Dim Condition As IQueryCommand
    Set Condition = mConditions(1)
    Dim Result As DataSet
    Set Result = Condition.Execute(Context)

    ' Combine each subsequent condition with the accumulated match.
    Dim Index As Long
    Dim ConditionResult As DataSet

    For Index = 2 To mConditions.Count
        Set Condition       = mConditions(Index)
        Set ConditionResult = Condition.Execute(Context)
        Set Result          = CombineMatches(Result, ConditionResult)
    Next Index

    Set IQueryCommand_Execute = Result
End Function

Private Function CombineMatches(ByVal LeftMatches As DataSet, ByVal RightMatches As DataSet) As DataSet

    If LeftMatches.RowCount <> RightMatches.RowCount Then
        Err.Raise vbObjectError + 1463, _
                  "QueryWhereGroupCommand.CombineMatches", _
                  "Match datasets must have the same row count."
    End If

    Dim Entries As New Collection

    Dim LeftField  As Field : Set LeftField  = LeftMatches.FieldByName("Match")
    Dim RightField As Field : Set RightField = RightMatches.FieldByName("Match")

    Dim RowIndex As Long
    For RowIndex = 1 To LeftField.Count
        Dim LeftEntry  As DataEntry : Set LeftEntry  = LeftField.Item(RowIndex)
        Dim RightEntry As DataEntry : Set RightEntry = RightField.Item(RowIndex)

        Call Entries.Add(DataEntry.Create(ApplyOperator(LeftEntry.Value, RightEntry.Value)))
    Next RowIndex

    Dim Definition As FieldDefinition
    Set Definition = FieldDefinition.Create("Match", vbBoolean, Validator.CreateFromCoreValueType(CoreBoolean))

    Dim ResultField As Field
    Set ResultField = Field.Create(Definition, Entries)

    Dim Result As DataSet
    Set Result = DataSet.Create()
    Call Result.AddField(ResultField)

    Set CombineMatches = Result
End Function

Private Function ApplyOperator(ByVal LeftValue As Boolean, ByVal RightValue As Boolean) As Boolean
    Select Case mLogicalOperator
        Case Andd  : ApplyOperator =      LeftValue And RightValue
        Case Orr   : ApplyOperator =      LeftValue Or  RightValue
        Case NAndd : ApplyOperator = Not (LeftValue And RightValue)
        Case NOrr  : ApplyOperator = Not (LeftValue Or  RightValue)
        Case XOrr  : ApplyOperator =     (LeftValue Xor RightValue)
        Case XNorr : ApplyOperator =     (LeftValue =   RightValue)
    End Select
End Function
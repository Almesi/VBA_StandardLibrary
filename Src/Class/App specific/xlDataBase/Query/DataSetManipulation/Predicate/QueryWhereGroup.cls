VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryWhereGroup"
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
' CREATION
'===============================================================================

Public Function Create(ByVal Operator As WhereGroupOperator) As QueryWhereGroup
    Set Create = New QueryWhereGroup
    Call Create.Initialize(Operator)
End Function

Friend Sub Initialize(ByVal Operator As WhereGroupOperator)
    Let mLogicalOperator = Operator
    Set mConditions      = New Collection
End Sub

Public Function AddCondition(ByVal Command As IQueryCommand) As QueryWhereGroup
    Call mConditions.Add(Command)
    Set AddCondition = Me
End Function

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If Not Context.HasCurrent Then Err.Raise vbObjectError + 1453, "QueryWhereGroup.Execute", "QueryContext has no current DataSet."
    If mConditions Is Nothing Then Err.Raise vbObjectError + 1461, "QueryWhereGroup.Execute", "Conditions have not been initialized."
    If mConditions.Count = 0  Then Err.Raise vbObjectError + 1461, "QueryWhereGroup.Execute", "At least one condition is required."

    Dim Source As QueryValue
    Set Source = Context.Current

    ' Evaluate the first condition to seed the accumulated match.
    Dim Condition As IQueryCommand
    Set Condition = mConditions(1)
    Dim Result As QueryValue
    Set Result = Condition.Execute(Context)

    ' Combine each subsequent condition with the accumulated match.
    Dim Index As Long
    Dim ConditionResult As QueryValue

    For Index = 2 To mConditions.Count
        Set Condition       = mConditions(Index)
        Set ConditionResult = Condition.Execute(Context)
        Set Result          = CombineMatches(Result.Value, ConditionResult.Value)
    Next Index

    Set IQueryCommand_Execute = Result
End Function

Private Function CombineMatches(ByVal LeftMatches As Collection, ByVal RightMatches As Collection) As QueryValue
    If LeftMatches.Count <> RightMatches.Count Then
        Err.Raise vbObjectError + 1463, _
                  "QueryWhereGroup.CombineMatches", _
                  "Match datasets must have the same row count."
    End If

    Dim Result As New Collection

    Dim RowIndex As Long
    For RowIndex = 1 To LeftMatches.Count
        Dim LeftValue  As Boolean : LeftValue  = LeftMatches.Item(RowIndex).Value
        Dim RightValue As Boolean : RightValue = RightMatches.Item(RowIndex).Value

        Call Result.Add(DataEntry.Create(ApplyOperator(LeftValue, RightValue)))
    Next RowIndex

    Set CombineMatches = QueryValue.CreateNormal(QueryValueKindSequence, Result)
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

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

Private mConditions      As Collection
Private mLogicalOperator As String


'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal LogicalOperator As String) As QueryWhereGroupCommand
    Set Create = New QueryWhereGroupCommand
    Call Create.Initialize(LogicalOperator)
End Function

Friend Sub Initialize(ByVal LogicalOperator As String)
    LogicalOperator = UCase(Trim(LogicalOperator))

    If LogicalOperator <> "AND" And LogicalOperator <> "OR" Then
        Err.Raise vbObjectError + 1420, _
                  "QueryWhereGroupCommand.Initialize", _
                  "Logical operator must be AND or OR."
    End If

    mLogicalOperator = LogicalOperator
    Set mConditions = New Collection
End Sub


'===============================================================================
' PROPERTIES
'===============================================================================

Public Property Get LogicalOperator() As String
    LogicalOperator = mLogicalOperator
End Property

Public Property Get ConditionCount() As Long
    ConditionCount = mConditions.Count
End Property


'===============================================================================
' ADD CONDITION
'===============================================================================

Public Function AddCondition(ByVal FieldName As String, ByVal Operator As String, ByVal Value As Variant) As QueryWhereGroupCommand
    Dim Condition As QueryWhereCommand
    Set Condition = QueryWhereCommand.Create(FieldName, Operator, Value)

    Call mConditions.Add(Condition)
    Set AddCondition = Me
End Function


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source As DataSet

    If Not Context.HasCurrent Then Err.Raise vbObjectError + 1421, "QueryWhereGroupCommand.Execute", "QueryContext has no current DataSet."
    If mConditions.Count = 0  Then Err.Raise vbObjectError + 1422, "QueryWhereGroupCommand.Execute", "At least one condition is required."

    Set Source = Context.Current
    Call ValidateConditions(Source)

    Set IQueryCommand_Execute = FilterDataSet(Source)
End Function


'===============================================================================
' VALIDATE CONDITIONS
'===============================================================================

Private Sub ValidateConditions(ByVal Source As DataSet)
    Dim Condition As QueryWhereCommand
    Dim Field As Field

    For Each Condition In mConditions
        Set Field = Source.FieldByName(Condition.FieldName)

        If Field Is Nothing Then
            Err.Raise vbObjectError + 1423, "QueryWhereGroupCommand.ValidateConditions", "Field '" & Condition.FieldName & "' was not found."
        End If
    Next Condition
End Sub


'===============================================================================
' FILTER
'===============================================================================

Private Function FilterDataSet(ByVal Source As DataSet) As DataSet
    Dim Result As DataSet
    Dim RowIndex As Long

    Set Result = New DataSet

    Call CopyStructure(Source, Result)

    For RowIndex = 1 To Source.RowCount
        If Matches(Source, RowIndex) Then
            Call Result.AddRow(Source.Row(RowIndex))
        End If
    Next RowIndex

    Set FilterDataSet = Result
End Function


'===============================================================================
' MATCH
'===============================================================================

Private Function Matches(ByVal Source As DataSet, ByVal RowIndex As Long) As Boolean
    Dim Condition       As QueryWhereCommand
    Dim FieldIndex      As Long
    Dim CurrentValue    As Variant
    Dim ConditionResult As Boolean

    If mLogicalOperator = "AND" Then
        Matches = True
    Else
        Matches = False
    End If

    For Each Condition In mConditions
        FieldIndex = FieldIndexByName(Source, Condition.FieldName)
        CurrentValue = Source.DataEntryAt(RowIndex, FieldIndex).Value
        ConditionResult = Compare(CurrentValue, Condition.Operator, Condition.Value)

        If mLogicalOperator = "AND" Then
            If Not ConditionResult Then
                Matches = False
                Exit Function
            End If
        Else
            If ConditionResult Then
                Matches = True
                Exit Function
            End If
        End If

    Next Condition
End Function


'===============================================================================
' COMPARE
'===============================================================================

Private Function Compare(ByVal LeftValue As Variant, ByVal Operator As String, ByVal RightValue As Variant) As Boolean
    Select Case Operator
        Case "="           : Compare = (LeftValue = RightValue)
        Case "<>"          : Compare = (LeftValue <> RightValue)
        Case ">"           : Compare = (LeftValue > RightValue)
        Case ">="          : Compare = (LeftValue >= RightValue)
        Case "<"           : Compare = (LeftValue < RightValue)
        Case "<="          : Compare = (LeftValue <= RightValue)
        Case "LIKE"        : Compare = (InStr(1, CStr(LeftValue), CStr(RightValue), vbTextCompare) > 0)
        Case "IS NULL"     : Compare = IsNull(LeftValue) Or IsEmpty(LeftValue)
        Case "IS NOT NULL" : Compare = Not (IsNull(LeftValue) Or IsEmpty(LeftValue))
        Case Else
            Err.Raise vbObjectError + 1424, _
                      "QueryWhereGroupCommand.Compare", _
                      "Unsupported operator '" & Operator & "'."
    End Select
End Function


'===============================================================================
' FIELD INDEX
'===============================================================================

Private Function FieldIndexByName(ByVal Source As DataSet, ByVal Name As String) As Long
    Dim Index As Long

    For Index = 1 To Source.FieldCount
        If StrComp(Source.FieldDefinition(Index).Name, Name, vbTextCompare) = 0 Then

            FieldIndexByName = Index
            Exit Function
        End If
    Next Index

    Err.Raise vbObjectError + 1425, "QueryWhereGroupCommand.FieldIndexByName", "Field '" & Name & "' was not found."
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
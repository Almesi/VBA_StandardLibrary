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
' PROPERTIES
'===============================================================================

Public Property Get FieldName() As String
    FieldName = mFieldName
End Property


Public Property Get Operator() As String
    Operator = mOperator
End Property


Public Property Get Value() As Variant
    Value = mValue
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source As DataSet
    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1412, _
                  "QueryWhereCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set Source = Context.Current
    Call ValidateField(Source)
    Set IQueryCommand_Execute = FilterDataSet(Source)
End Function


'===============================================================================
' VALIDATE FIELD
'===============================================================================

Private Sub ValidateField(ByVal Source As DataSet)
    Dim Field As Field
    Set Field = Source.FieldByName(mFieldName)
    If Field Is Nothing Then
        Err.Raise vbObjectError + 1413, _
                  "QueryWhereCommand.Execute", _
                  "Field '" & mFieldName & "' was not found."
    End If
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
    Dim FieldIndex As Long
    Dim Field As Field
    Dim CurrentValue As Variant

    Set Field = Source.FieldByName(mFieldName)
    FieldIndex = FieldIndexByName(Source, mFieldName)
    CurrentValue = Source.DataEntryAt(RowIndex, FieldIndex).Value
    Matches = Compare(CurrentValue, mOperator, mValue)
End Function


'===============================================================================
' COMPARE
'===============================================================================

Private Function Compare(ByVal LeftValue As Variant, ByVal Operator As String, ByVal RightValue As Variant) As Boolean
    Select Case Operator
        Case "="           : Compare = (LeftValue =  RightValue)
        Case "<>"          : Compare = (LeftValue <> RightValue)
        Case ">"           : Compare = (LeftValue >  RightValue)
        Case ">="          : Compare = (LeftValue >= RightValue)
        Case "<"           : Compare = (LeftValue <  RightValue)
        Case "<="          : Compare = (LeftValue <= RightValue)
        Case "LIKE"        : Compare = (InStr(1, CStr(LeftValue), CStr(RightValue), vbTextCompare) > 0)
        Case "IS NULL"     : Compare = IsNull(LeftValue) Or IsEmpty(LeftValue)
        Case "IS NOT NULL" : Compare = Not (IsNull(LeftValue) Or IsEmpty(LeftValue))
        Case Else
            Err.Raise vbObjectError + 1414, _
                      "QueryWhereCommand.Compare", _
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
Attribute VB_Name = "QueryHelpers"


Option Explicit

Public Function QueryAssert(ByVal Context As QueryContext, ByVal ExpectedKind As QueryValueKind, Optional ByVal FieldName As String = Empty) As Boolean
    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1453, "QueryAssert", "QueryContext has no current."
        Exit Function
    End If

    Dim Current As QueryValue
    Set Current = Context.Current
    If Not Current.Kind = ExpectedKind Then
        Err.Raise vbObjectError + 1453, "QueryAssert", _
                                        "QueryContext has wrong Kind of QueryValue. Expected: " & ExpectedKind & " Actual: " & Current.Kind
    End If

    If FieldName <> Empty And Current.Kind = QueryValueKindDataSet Then
        Dim Field As Field
        Set Field = Current.Value.FieldByName(FieldName)
        If Field Is Nothing Then
            Err.Raise vbObjectError + 1454, "QueryAssert", "Field '" & FieldName & "' was not found."
            Exit Function
        End If
    End If

    QueryAssert = True
End Function
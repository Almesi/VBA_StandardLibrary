VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryOrderByCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mFields As Collection


'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ParamArray Fields() As Variant) As QueryOrderByCommand
    Set Create = New QueryOrderByCommand
    Call Create.Initialize(Fields)
End Function


Friend Sub Initialize(ByRef Fields() As Variant)
    Dim Index As Long
    Dim Item As Variant
    Dim Field As QueryOrderByField

    Set mFields = New Collection

    For Index = LBound(Fields) To UBound(Fields)
        Set Field = Fields(Index)

        If Field Is Nothing Then
            Err.Raise vbObjectError + 1431, _
                      "QueryOrderByCommand.Initialize", _
                      "OrderBy field cannot be Nothing."
        End If

        Call mFields.Add(Field)
    Next Index

    If mFields.Count = 0 Then
        Err.Raise vbObjectError + 1432, _
                  "QueryOrderByCommand.Initialize", _
                  "At least one OrderBy field is required."
    End If
End Sub


'===============================================================================
' PROPERTIES
'===============================================================================

Public Property Get Fields() As Collection
    Set Fields = mFields
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source As DataSet

    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1433, _
                  "QueryOrderByCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set Source = Context.Current
    Call ValidateFields(Source)

    Set IQueryCommand_Execute = SortDataSet(Source)
End Function


'===============================================================================
' VALIDATE FIELDS
'===============================================================================

Private Sub ValidateFields(ByVal Source As DataSet)
    Dim Item As Variant
    Dim Field As QueryOrderByField
    Dim SourceField As Field

    For Each Item In mFields
        Set Field = Item
        Set SourceField = Source.FieldByName(Field.FieldName)

        If SourceField Is Nothing Then
            Err.Raise vbObjectError + 1434, _
                      "QueryOrderByCommand.Execute", _
                      "Field '" & Field.FieldName & "' was not found."
        End If
    Next Item
End Sub


'===============================================================================
' SORT
'===============================================================================

Private Function SortDataSet(ByVal Source As DataSet) As DataSet
    Dim Result As DataSet
    Dim RowIndexes() As Long
    Dim Index As Long

    Set Result = New DataSet
    Call CopyStructure(Source, Result)

    If Source.RowCount = 0 Then
        Set SortDataSet = Result
        Exit Function
    End If

    ReDim RowIndexes(1 To Source.RowCount)

    For Index = 1 To Source.RowCount
        RowIndexes(Index) = Index
    Next Index

    Call QuickSort(Source, RowIndexes, LBound(RowIndexes), UBound(RowIndexes))

    For Index = LBound(RowIndexes) To UBound(RowIndexes)
        Call Result.AddRow(Source.Row(RowIndexes(Index)))
    Next Index

    Set SortDataSet = Result
End Function


'===============================================================================
' QUICK SORT
'===============================================================================

Private Sub QuickSort(ByVal Source As DataSet, ByRef RowIndexes() As Long, ByVal First As Long, ByVal Last As Long)
    Dim Low As Long
    Dim High As Long
    Dim Pivot As Long
    Dim Temp As Long

    Low = First
    High = Last
    Pivot = RowIndexes((First + Last) \ 2)

    Do While Low <= High

        Do While CompareRows(Source, RowIndexes(Low), Pivot) < 0
            Low = Low + 1
        Loop

        Do While CompareRows(Source, RowIndexes(High), Pivot) > 0
            High = High - 1
        Loop

        If Low <= High Then
            Temp = RowIndexes(Low)
            RowIndexes(Low) = RowIndexes(High)
            RowIndexes(High) = Temp

            Low = Low + 1
            High = High - 1
        End If
    Loop

    If First < High Then
        Call QuickSort(Source, RowIndexes, First, High)
    End If

    If Low < Last Then
        Call QuickSort(Source, RowIndexes, Low, Last)
    End If
End Sub


'===============================================================================
' COMPARE ROWS
'===============================================================================

Private Function CompareRows(ByVal Source As DataSet, ByVal LeftRow As Long, ByVal RightRow As Long) As Long
    Dim Item As Variant
    Dim OrderField As QueryOrderByField
    Dim FieldIndex As Long
    Dim LeftValue As Variant
    Dim RightValue As Variant
    Dim Comparison As Long

    For Each Item In mFields
        Set OrderField = Item

        FieldIndex = FieldIndexByName(Source, OrderField.FieldName)

        LeftValue = Source.DataEntryAt(LeftRow, FieldIndex).Value
        RightValue = Source.DataEntryAt(RightRow, FieldIndex).Value

        Comparison = CompareValues(LeftValue, RightValue)

        If Comparison <> 0 Then
            If OrderField.Descending Then
                Comparison = -Comparison
            End If

            CompareRows = Comparison
            Exit Function
        End If
    Next Item

    CompareRows = 0
End Function


'===============================================================================
' COMPARE VALUES
'===============================================================================

Private Function CompareValues(ByVal LeftValue As Variant, ByVal RightValue As Variant) As Long
    If IsNull(LeftValue) Or IsEmpty(LeftValue) Then
        If IsNull(RightValue) Or IsEmpty(RightValue) Then
            CompareValues = 0
        Else
            CompareValues = -1
        End If
        Exit Function
    End If

    If IsNull(RightValue) Or IsEmpty(RightValue) Then
        CompareValues = 1
        Exit Function
    End If

    If LeftValue < RightValue Then
        CompareValues = -1
    ElseIf LeftValue > RightValue Then
        CompareValues = 1
    Else
        CompareValues = 0
    End If
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
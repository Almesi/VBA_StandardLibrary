VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryOrderBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mFields As Collection


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ParamArray Fields() As Variant) As QueryOrderBy
    Set Create = New QueryOrderBy
    Call Create.Initialize(Fields)
End Function


Friend Sub Initialize(ByRef Fields() As Variant)
    Set mFields = Fields
End Sub

'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet) = False Then Exit Function

    Dim Source As DataSet
    Set Source = Context.Current.Value
    Call ValidateFields(Source)

    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, SortDataSet(Source))
End Function


Private Sub ValidateFields(ByVal Source As DataSet)
    Dim Item As Variant
    Dim Field As QueryOrderByField
    Dim SourceField As Field

    For Each Item In mFields
        Set Field = Item
        Set SourceField = Source.FieldByName(Field.FieldName)

        If SourceField Is Nothing Then
            Err.Raise vbObjectError + 1434, _
                      "QueryOrderBy.Execute", _
                      "Field '" & Field.FieldName & "' was not found."
        End If
    Next Item
End Sub

Private Function SortDataSet(ByVal Source As DataSet) As DataSet
    Dim Result       As DataSet
    Dim RowIndexes() As Long
    Dim Index        As Long

    Set Result = Source.CopyStructure

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


Private Sub QuickSort(ByVal Source As DataSet, ByRef RowIndexes() As Long, ByVal First As Long, ByVal Last As Long)
    Dim Low   As Long
    Dim High  As Long
    Dim Pivot As Long
    Dim Temp  As Long

    Low   = First
    High  = Last
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


Private Function CompareRows(ByVal Source As DataSet, ByVal LeftRow As Long, ByVal RightRow As Long) As Long
    Dim Item       As Variant
    Dim OrderField As QueryOrderByField
    Dim Field      As Field
    Dim LeftValue  As Variant
    Dim RightValue As Variant
    Dim Comparison As Long

    For Each Item In mFields
        Set OrderField = Item

        Set Field  = Source.FieldByName(OrderField.FieldName)
        LeftValue  = Field.Item(LeftRow).Value
        RightValue = Field.Item(RightRow).Value
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
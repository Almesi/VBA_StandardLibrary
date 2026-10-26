VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryBottom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mCount As Long


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal Count As Long) As QueryBottom
    Set Create = New QueryBottom
    Call Create.Initialize(Count)
End Function


Friend Sub Initialize(ByVal Count As Long)
    If Count < 0 Then
        Err.Raise vbObjectError + 1450, _
                  "QueryBottom.Initialize", _
                  "Count cannot be negative."
    End If

    Let mCount = Count
End Sub


'===============================================================================
' PROPERTIES
'===============================================================================

Public Property Get Count() As Long
    Count = mCount
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    If QueryAssert(Context, QueryValueKindDataSet) = False Then Exit Function

    Dim Source As DataSet
    Set Source = Context.Current.Value

    Dim Result As DataSet
    Set Result = Source.CopyStructure

    If mCount = 0 Then
        Set IQueryCommand_Execute = Result
        Exit Function
    End If

    Dim StartIndex As Long
    StartIndex = Source.RowCount - mCount + 1
    If StartIndex < 1 Then
        StartIndex = 1
    End If

    Dim RowIndex As Long
    For RowIndex = StartIndex To Source.RowCount
        Call Result.AddRow(Source.Row(RowIndex))
    Next RowIndex

    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, Result)
End Function
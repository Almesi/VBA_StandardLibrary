VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryUnionCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mDataSets As Collection


'===============================================================================
' INITIALIZATION
'===============================================================================

Private Sub Class_Initialize()
    Set mDataSets = New Collection
End Sub

Public Function Create(ParamArray Names() As Variant) As QueryUnionCommand
    Set Create = New QueryUnionCommand

    Dim i As Long

    For i = 0 To UBound(Names)
        Call Create.Add(CStr(Names(i)))
    Next i
End Function


Public Sub Add(ByVal Name As String)
    If Len(Trim(Name)) = 0 Then Exit Sub
    Call mDataSets.Add(Name)
End Sub


Public Property Get DataSets() As Collection
    Set DataSets = mDataSets
End Property


'===============================================================================
' EXECUTE
'===============================================================================
Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Result As DataSet
    Dim Source As DataSet
    Dim DataSetName As Variant

    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1403, _
                  "QueryUnionCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set Result = Context.Current

    For Each DataSetName In mDataSets
        Set Source = Context.DataSetByName(CStr(DataSetName))

        If Source Is Nothing Then
            Err.Raise vbObjectError + 1404, _
                      "QueryUnionCommand.Execute", _
                      "DataSet '" & CStr(DataSetName) & "' was not found."
        End If

        Set Result = DataSet.Merge(Result, Source)
    Next DataSetName
    Set IQueryCommand_Execute = Result
End Function
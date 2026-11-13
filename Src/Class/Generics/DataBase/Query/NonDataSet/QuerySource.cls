VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QuerySource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mSourceName As String


'===============================================================================
' CREATION
'===============================================================================



Public Function Create(ByVal SourceName As String) As QuerySource
    Set Create = New QuerySource
    Call Create.Initialize(SourceName)
End Function

Friend Sub Initialize(ByVal SourceName As String)
    Let mSourceName = SourceName
End Sub


'===============================================================================
' SOURCE
'===============================================================================

Public Property Get SourceName() As String
    SourceName = mSourceName
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As QueryValue
    Dim Result As DataSet
    Set Result = Context.DataSetByName(mSourceName)
    If Result Is Nothing Then
        Err.Raise vbObjectError + 1402, _
                  "QuerySource.Execute", _
                  "Query source '" & mSourceName & "' was not found."
    End If
    Set IQueryCommand_Execute = QueryValue.CreateNormal(QueryValueKindDataSet, Result)
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QuerySourceCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mSourceName As String


'===============================================================================
' INITIALIZATION
'===============================================================================



Public Function Create(ByVal SourceName As String) As QuerySourceCommand
    Set Create = New QuerySourceCommand
    Call Create.Initialize(SourceName)
End Function

Friend Sub Initialize(ByVal SourceName As String)
    If Len(Trim(SourceName)) = 0 Then

        Err.Raise vbObjectError + 1401, _
                  "QuerySourceCommand.Initialize", _
                  "Source name cannot be empty."
    End If
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

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Result As DataSet
    Set Result = Context.DataSetByName(mSourceName)
    If Result Is Nothing Then
        Err.Raise vbObjectError + 1402, _
                  "QuerySourceCommand.Execute", _
                  "Query source '" & mSourceName & "' was not found."
    End If
    Set IQueryCommand_Execute = Result
End Function
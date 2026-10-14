VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryAddCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mAddName As String


'===============================================================================
' INITIALIZATION
'===============================================================================



Public Function Create(ByVal AddName As String) As QueryAddCommand
    Set Create = New QueryAddCommand
    Call Create.Initialize(AddName)
End Function

Friend Sub Initialize(ByVal AddName As String)
    If Len(Trim(AddName)) = 0 Then

        Err.Raise vbObjectError + 1401, _
                  "QueryAddCommand.Initialize", _
                  "Add name cannot be empty."
    End If
    Let mAddName = AddName
End Sub


'===============================================================================
' Add
'===============================================================================

Public Property Get AddName() As String
    AddName = mAddName
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Result As DataSet
    Set Result = Context.Current
    If Result Is Nothing Then
        Err.Raise vbObjectError + 1402, _
                  "QueryAddCommand.Execute", _
                  "Query Add '" & mAddName & "' was not found."
    End If
    Call Context.Add(NamedDataSet.Create(AddName, Result))
    Set IQueryCommand_Execute = Result
End Function
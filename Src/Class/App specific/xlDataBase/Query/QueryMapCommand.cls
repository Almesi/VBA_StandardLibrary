VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryMapCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IQueryCommand

Private mFieldName  As String
Private mFunc       As IFunction


'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal nFieldName As String, ByVal nFunc As IFunction) As QueryMapCommand
    Set Create = New QueryMapCommand
    Call Create.Initialize(nFieldName, nFunc)
End Function

Friend Sub Initialize(ByVal nFieldName As String, ByVal nFunc As IFunction)
    Let mFieldName = FieldName
    Set mFunc      = nFunc
End Sub


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source As DataSet

    If Not Context.HasCurrent Then Err.Raise vbObjectError + 1422, "QueryMapCommand.Execute", "QueryContext has no current DataSet."

    Set Source = Context.Current
    Set IQueryCommand_Execute = MapDataSet(Source)
End Function

'===============================================================================
' MAP
'===============================================================================

Private Function MapDataSet(ByVal Source As DataSet) As DataSet
    Dim Result      As DataSet
    Dim Field       As Field
    Dim Entries     As Collection
    Dim Entry       As DataEntry

    Set Result = Source.Copy()

    Set Field   = Result.FieldByName(mFieldName)
    Set Entries = Field.DataEntries
    For Each Entry In Entries
        Entry.Value = mFunc.Evaluate(Entry.Value)
    Next Entry

    Set MapDataSet = Result
End Function
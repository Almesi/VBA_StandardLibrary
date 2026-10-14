VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QuerySelectCommand"
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

Private Sub Class_Initialize()
    Set mFields = New Collection
End Sub

Public Function Create(ParamArray Names() As Variant) As QuerySelectCommand
    Set Create = New QuerySelectCommand
    Dim i As Long
    For i = 0 To Ubound(Names)
        Call Create.Add(Names(i))
    Next i
End Function


Public Sub Add(ByVal Name As String)
    If Len(Trim(Name)) = 0 Then Exit Sub
    Call mFields.Add(Name)
End Sub


Public Property Get Fields() As Collection
    Set Fields = mFields
End Property


'===============================================================================
' EXECUTE
'===============================================================================

Private Function IQueryCommand_Execute(ByVal Context As QueryContext) As DataSet
    Dim Source      As DataSet
    Dim Result      As DataSet
    Dim FieldName   As Variant
    Dim SourceField As Field
    Dim NewField    As Field
    
    If Not Context.HasCurrent Then
        Err.Raise vbObjectError + 1403, _
                  "QuerySelectCommand.Execute", _
                  "QueryContext has no current DataSet."
    End If

    Set Source = Context.Current
    Set Result = New DataSet

    For Each FieldName In mFields
        Set SourceField = Source.FieldByName(CStr(FieldName))
        If SourceField Is Nothing Then
            Err.Raise vbObjectError + 1404, _
                      "QuerySelectCommand.Execute", _
                      "Field '" & CStr(FieldName) & "' was not found."
        End If

        Set NewField = CopyField(SourceField)
        Call Result.AddField(NewField)
    Next FieldName
    Set IQueryCommand_Execute = Result
End Function


'===============================================================================
' COPY FIELD
'===============================================================================

Private Function CopyField(ByVal Source As Field) As Field
    Dim Result As Field
    Dim Entry As DataEntry

    Set Result = New Field

    Set Result.Definition = Source.Definition
    For Each Entry In Source.DataEntries
        Call Result.AddDataEntry(Entry)
    Next Entry
    Set CopyField = Result
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ChoiceCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mChoices As Object

'============================================================
'====================[Creation]==============================
'============================================================

Private Sub Class_Initialize()
    Set mChoices = CreateObject("Scripting.Dictionary")
    mChoices.CompareMode = vbTextCompare
End Sub


Public Function Create(ByVal ChoicesDataSet As DataSet, ByVal ChoiceValuesDataSet As DataSet) As ChoiceCollector
    Set Create = New ChoiceCollector
    Call Create.Initialize(ChoicesDataSet, ChoiceValuesDataSet)
End Function

Friend Sub Initialize(ByVal ChoicesDataSet As DataSet, ByVal ChoiceValuesDataSet As DataSet)
    Dim Row        As RowEntry
    Dim ChoiceID   As Long
    Dim ChoiceName As String
    Dim Value      As String

    For Each Row In ChoicesDataSet.Rows
        ChoiceID   = CLng(Row.ByName("ID").DataEntry.Value)
        ChoiceName = CStr(Row.ByName("Name").DataEntry.Value)

        If Not mChoices.Exists(ChoiceName) Then
            Call mChoices.Add(ChoiceName, New Collection)
        End If
    Next Row

    For Each Row In ChoiceValuesDataSet.Rows
        ChoiceID   = CLng(Row.ByName("DB_META_Choices::ID").DataEntry.Value)
        Value    = CStr(Row.ByName("Code").DataEntry.Value)
        ChoiceName = GetChoiceNameByID(ChoicesDataSet, ChoiceID)

        If Len(ChoiceName) > 0 Then
            Call mChoices(ChoiceName).Add(Value)
        End If
    Next Row
End Sub

'============================================================
'====================[Getters]===============================
'============================================================

Public Function GetChoices(ByVal ChoiceName As String) As Collection
    If Not mChoices.Exists(ChoiceName) Then
        Err.Raise _
            vbObjectError + 1000, _
            "ChoiceCollector.GetChoices", _
            "Unknown Choice: '" & ChoiceName & "'"
    End If
    Set GetChoices = mChoices(ChoiceName)
End Function


'============================================================
'====================[Functions]=============================
'============================================================

Public Function Exists(ByVal ChoiceName As String) As Boolean
    Exists = mChoices.Exists(ChoiceName)
End Function

'============================================================
'====================[Helpers]===============================
'============================================================

Private Function GetChoiceNameByID(ByVal ChoicesDataSet As DataSet, ByVal ChoiceID As Long) As String
    Dim Row As RowEntry

    For Each Row In ChoicesDataSet.Rows
        If CLng(Row.ByName("ID").DataEntry.Value) = ChoiceID Then
            GetChoiceNameByID = CStr(Row.ByName("Name").DataEntry.Value)
            Exit Function
        End If
    Next Row
    GetChoiceNameByID = vbNullString
End Function
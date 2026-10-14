VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mSource     As ListObject
Private mClassifier As FieldClassifier
Private mDataSet    As DataSet

'============================================================
'====================[Creation]==============================
'============================================================

Private Sub Class_Initialize()
    Set mDataSet = New DataSet
End Sub

Public Function CreateFromTable(ByVal SourceTable As ListObject, Optional ByVal Classifier As FieldClassifier = Nothing) As Table
    Set CreateFromTable = New Table

    Call CreateFromTable.Initialize(SourceTable, Classifier)
    Call CreateFromTable.Refresh()
End Function

Friend Sub Initialize(ByVal Source As ListObject, Optional ByVal Classifier As FieldClassifier = Nothing)
    Set mSource     = Source
    Set mClassifier = Classifier
End Sub

'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Get DataSet() As DataSet
    Set DataSet = mDataSet
End Property

Public Property Set DataSet(ByVal Value As DataSet)
    Set mDataSet = Value
End Property

'============================================================
'====================[Functions]=============================
'============================================================

Public Sub Refresh()
    Call DataSet.Clear()
    Call CreateFieldDefinitions()
    Call CreateRows()
End Sub

Private Sub CreateFieldDefinitions()
    Dim FieldDefinition As FieldDefinition
    Dim FieldIndex As Long
    Dim Name As String

    For FieldIndex = 1 To mSource.ListColumns.Count
        Name = mSource.ListColumns(FieldIndex).Name

        If mClassifier Is Nothing Then
            Set FieldDefinition = New FieldDefinition
            FieldDefinition.Name = Name
            FieldDefinition.DataType = vbVariant
            Set FieldDefinition.Validator = Validator.Create(GenericChecker.Create(vbVariant), GenericConverter.Create(vbVariant))
        Else
            Set FieldDefinition = mClassifier.Classify(Name)
        End If

        Call DataSet.AddFieldDefinition(FieldDefinition)
    Next FieldIndex
End Sub

'============================================================
'====================[Helpers]===============================
'============================================================

Private Sub CreateRows()
    Dim RowIndex As Long
    Dim Row As RowEntry

    If mSource.DataBodyRange Is Nothing Then Exit Sub

    For RowIndex = 1 To mSource.ListRows.Count
        Set Row = CreateRow(RowIndex)
        Call DataSet.AddRow(Row)
    Next RowIndex
End Sub

Private Function CreateRow(ByVal RowIndex As Long) As RowEntry
    Dim Result     As RowEntry
    Dim Entry      As FieldEntry
    Dim DataCell   As Range
    Dim FieldIndex As Long

    Set Result = New RowEntry

    For FieldIndex = 1 To mSource.ListColumns.Count
        Set DataCell = mSource.DataBodyRange.Cells(RowIndex, FieldIndex)

        Set Entry = FieldEntry.Create(DataSet.FieldDefinition(FieldIndex), DataEntry.Create(DataCell.Value))

        Call Result.AddFieldEntry(Entry)
    Next FieldIndex

    Set CreateRow = Result
End Function
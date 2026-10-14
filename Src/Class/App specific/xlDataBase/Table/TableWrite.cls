VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "TableWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mTable As Table

'============================================================
'====================[Creation]==============================
'============================================================

Public Function CreateFromTable(ByVal SourceTable As ListObject, Optional ByVal Classifier As FieldClassifier = Nothing) As TableWrite
    Dim CreateFromTable As TableWrite
    Set CreateFromTable = New TableWrite
    Call CreateFromTable.Initialize(SourceTable, Classifier)
End Function

Friend Sub Initialize(ByVal SourceTable As ListObject, Optional ByVal Classifier As FieldClassifier = Nothing)
    Set mTable = Table.Create(SourceTable, Classifier)
End Sub

'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Set Table(ByVal Value As Table)
    Set mTable = Value
End Property

Public Property Get Table() As Table
    Set Table = mTable
End Property

'============================================================
'====================[Functions]=============================
'============================================================

Public Sub WriteToTable(ByVal DestinationTable As ListObject)
    Dim DataCell   As Range
    Dim RowIndex   As Long
    Dim FieldIndex As Long

    If DestinationTable.DataBodyRange Is Nothing Then Exit Sub
    For RowIndex = 1 To Table.RowCount
        For FieldIndex = 1 To Table.FieldDefinitionCount
            Set DataCell = DestinationTable.DataBodyRange.Cells(RowIndex, FieldIndex)
            DataCell.Value = Table.DataEntryAt(RowIndex, FieldIndex).Value
        Next FieldIndex
    Next RowIndex
End Sub

'============================================================
'====================[Helpers]===============================
'============================================================
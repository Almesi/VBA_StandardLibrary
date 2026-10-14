VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "RowEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mFieldEntries As Collection

'============================================================
'====================[Creation]==============================
'============================================================

Private Sub Class_Initialize()
    Set mFieldEntries = New Collection
End Sub

Public Function Create(ByVal nFieldEntries As Collection) As RowEntry
    Set Create = New RowEntry

    Dim Entry As FieldEntry
    For Each Entry In nFieldEntries
        Create.AddFieldEntry(Entry)
    Next Entry
End Function

'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Get FieldEntries() As Collection
    Set FieldEntries = mFieldEntries
End Property

Public Property Get Count() As Long
    Count = mFieldEntries.Count
End Property

Public Property Get Item(ByVal Index As Long) As FieldEntry
    Set Item = mFieldEntries(Index)
End Property

'============================================================
'====================[Functions]=============================
'============================================================

Public Function ByName(ByVal Name As String) As FieldEntry
    Dim i As Long
    For i = 1 To Count()
        If mFieldEntries(i).Definition.Name = Name Then
            Set ByName = mFieldEntries(i)
            Exit Function
        End If
    Next i
    Set Field = Nothing
End Function

Public Sub AddFieldEntry(ByVal Value As FieldEntry)
    Call mFieldEntries.Add(Value)
End Sub
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mDefinition  As FieldDefinition
Private mDataEntries As Collection

'============================================================
'====================[Creation]==============================
'============================================================

Private Sub Class_Initialize()
    Set mDataEntries = New Collection
End Sub

Public Function Create(ByVal nDefinition As FieldDefinition, ByVal nDataEntries As Collection) As Field
    Dim Create = New Field
    Create.Definition = nDefinition

    Dim Entry As DataEntry
    For Each Entry In nDataEntries
        Create.AddDataEntry(Entry)
    Next Entry
End Function

'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Get Definition()              As FieldDefinition  :    Set Definition  = mDefinition          : End Property
Public Property Set Definition(ByVal Value    As FieldDefinition) :    Set mDefinition = Value                : End Property
Public Property Get DataEntries()             As Collection       :    Set DataEntries = mDataEntries         : End Property
Public Property Get Count()                   As Long             :    Let Count       = mDataEntries.Count   : End Property
Public Property Get Item(ByVal Index As Long) As DataEntry        :    Set Item        = mDataEntries(Index)  : End Property

'============================================================
'====================[Functions]=============================
'============================================================

Public Sub AddDataEntry(ByVal Value As DataEntry)
    Call mDataEntries.Add(Value)
End Sub
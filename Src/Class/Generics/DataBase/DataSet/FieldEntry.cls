VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "FieldEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mDefinition As FieldDefinition
Private mDataEntry  As DataEntry

'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Set Definition(ByVal Value As FieldDefinition) : Set mDefinition  = Value       : End Property
Public Property Set DataEntry(ByVal  Value As DataEntry)       : Set mDataEntry   = Value       : End Property
Public Property Get Definition()           As FieldDefinition  : Set Definition   = mDefinition : End Property
Public Property Get DataEntry()            As DataEntry        : Set DataEntry    = mDataEntry  : End Property

'============================================================
'====================[Creation]==============================
'============================================================

Public Function Create(ByVal nDefinition As FieldDefinition, ByVal nDataEntry As DataEntry) As FieldEntry
    Set Create = New FieldEntry
    Set Create.Definition = nDefinition
    Set Create.DataEntry  = nDataEntry
End Function

'============================================================
'====================[Functions]=============================
'============================================================
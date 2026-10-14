VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldEntryModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Type TModel
    Entry As FieldEntry
End Type

Private This As TModel

Public Property Get Entry() As FieldEntry
    Set Entry = This.Entry
End Property

Public Property Set Entry(ByVal Entry As FieldEntry)
    Set This.Entry = Entry
End Property

Public Property Get Self() As FieldEntryModel
    Set Self = Me
End Property
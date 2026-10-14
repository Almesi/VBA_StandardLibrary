VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "RowModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TModel
    Row As RowEntry
End Type

Private This As TModel


Public Property Get Row() As RowEntry
    Set Row = This.Row
End Property

Public Property Set Row(ByVal Value As RowEntry)
    Set This.Row = Value
End Property

Public Property Get Self() As RowModel
    Set Self = Me
End Property
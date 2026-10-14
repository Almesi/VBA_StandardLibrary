VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "FieldModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TModel
    Field As Field
End Type

Private This As TModel


Public Property Get Field() As Field
    Set Field = This.Field
End Property

Public Property Set Field(ByVal Value As Field)
    Set This.Field = Value
End Property

Public Property Get Self() As FieldModel
    Set Self = Me
End Property
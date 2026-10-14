VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "ValueModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TModel
    Value As Variant
End Type

Private This As TModel

Public Property Set Value(ByVal Value As Variant)
    Set This.Value = Value
End Property
Public Property Let Value(ByVal Value As Variant)
    Let This.Value = Value
End Property

Public Property Get Value() As Variant
    If ISObject(This.Value) Then
        Set Value = This.Value
    Else
        Let Value = This.Value
    End If
End Property

Public Property Get Self() As ValueModel
    Set Self = Me
End Property
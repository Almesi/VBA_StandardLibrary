VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SingleInputModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private Type TModel
    Value As String
End Type

Private This As TModel

Public Property Get Value() As String
    Value = This.Value
End Property

Public Property Let Value(ByVal Value As String)
    This.Value = Value
End Property

Public Property Get Self() As SingleInputModel
    Set Self = Me
End Property
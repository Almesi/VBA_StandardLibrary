VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Implements IGetLetSet

Private mValue As Object

Public Property Get Value() As Variant
    Set Value = mValue
End Property

Public Property Set Value(ByVal NewValue As Object)
    Set mValue = NewValue
End Property

Private Property Get IGetLetSet_Value() As Variant
    Set IGetLetSet_Value = mValue
End Property

Private Property Let IGetLetSet_Value(ByVal NewValue As Variant)
    Err.Raise 438, "ObjectValue", "Use Set to assign an object reference."
End Property

Private Property Set IGetLetSet_Value(ByVal NewValue As Object)
    Set mValue = NewValue
End Property

Public Sub IGetLetSet_Assign(ByRef Destination, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Destination = Value
    Else
        Let Destination = Value
    End If
End Sub
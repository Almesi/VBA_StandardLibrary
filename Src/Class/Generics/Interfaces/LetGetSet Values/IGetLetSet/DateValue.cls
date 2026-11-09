VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Implements IGetLetSet

Private mValue As Date

Public Property Get Value() As Variant
    Value = mValue
End Property

Public Property Let Value(ByVal NewValue As Variant)
    If VarType(NewValue) <> vbDate Then
        Err.Raise 13, "DateValue", "Expected Date; received " & TypeName(NewValue)
    End If
    mValue = NewValue
End Property

Private Property Get IGetLetSet_Value() As Variant
    IGetLetSet_Value = mValue
End Property

Private Property Let IGetLetSet_Value(ByVal NewValue As Variant)
    If VarType(NewValue) <> vbDate Then
        Err.Raise 13, "DateValue", "Expected Date; received " & TypeName(NewValue)
    End If
    mValue = NewValue
End Property

Private Property Set IGetLetSet_Value(ByVal NewValue As Object)
    Err.Raise 438, "DateValue", "This value wrapper stores a scalar, not an object."
End Property

Public Sub IGetLetSet_Assign(ByRef Destination, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Destination = Value
    Else
        Let Destination = Value
    End If
End Sub
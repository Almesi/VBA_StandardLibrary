VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DoubleValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IGet
Implements ILet
Implements IGetLet

Private mValue As Double

Public Function Create(ByVal nValue As Double) As DoubleValue
    Set Create = New DoubleValue
    Create.Value = nValue
End Function

Private Sub Validate(ByVal NewValue As Variant)
    If VarType(NewValue) <> vbDouble Then Err.Raise 13, "DoubleValue", "Expected Double; received " & TypeName(NewValue)
End Sub

Public Property Let Value(ByVal NewValue As Variant) : Call Validate(NewValue) : Let mValue = NewValue : End Property
Public Property Get Value()              As Variant  :                         : Let Value  = mValue   : End Property

Private Property Let ILet_Value(ByVal       NewValue      As Variant) : Let Value         = NewValue : End Property
Private Property Let IGetLet_Value(ByVal    NewValue      As Variant) : Let Value         = NewValue : End Property
Private Property Get IGet_Value()                         As Variant  : Let IGet_Value    = Value    : End Property
Private Property Get IGetLet_Value()                      As Variant  : Let IGetLet_Value = Value    : End Property
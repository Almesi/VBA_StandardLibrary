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

Implements IGet
Implements ISet
Implements IGetSet

Private mValue As Object

Public Function Create(ByVal nValue As Object) As ObjectValue
    Set Create = New ObjectValue
    Set Create.Value = nValue
End Function

Private Sub Validate(ByVal NewValue As Variant)
    If VarType(NewValue) <> vbObject Then Err.Raise 13, "ObjectValue", "Expected Object; received " & TypeName(NewValue)
End Sub

Public Property Set Value(ByVal NewValue As Variant) : Call Validate(NewValue) : Set mValue = NewValue : End Property
Public Property Get Value()              As Variant  :                         : Set Value  = mValue   : End Property

Private Property Set ISet_Value(ByVal       NewValue      As Variant) : Set Value         = NewValue : End Property
Private Property Set IGetSet_Value(ByVal    NewValue      As Variant) : Set Value         = NewValue : End Property
Private Property Get IGet_Value()                         As Variant  : Set IGet_Value    = Value    : End Property
Private Property Get IGetSet_Value()                      As Variant  : Set IGetSet_Value = Value    : End Property
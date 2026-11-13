VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariantValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IGet
Implements ILet
Implements ISet
Implements IGetLet
Implements IGetSet
Implements IGetLetSet

Private mValue As Variant

Public Function Create(ByVal nValue As Variant) As VariantValue
    Set Create = New VariantValue
    Call Create.Assign(nValue)
End Function

Private Sub Validate(ByVal NewValue As Variant)
    '
End Sub

Public Property Let Value(ByVal NewValue As Variant) : Call Validate(NewValue) : Let mValue = NewValue : End Property
Public Property Set Value(ByVal NewValue As Variant) : Call Validate(NewValue) : Set mValue = NewValue : End Property
Public Property Get Value()              As Variant  :                         : Let Value  = mValue   : End Property

Public Sub Assign(ByVal Value As Variant)
    If IsObject(Value) Then
        Set mValue = Value
    Else
        Let mValue = Value
    End If
End Sub
Public Sub Extract(ByRef Destination As Variant)
    If IsObject(mValue) Then
        Set Destination = mValue
    Else
        Let Destination = mValue
    End If
End Sub

Private Property Let ILet_Value(ByVal       NewValue      As Variant) : Let Value = NewValue           : End Property
Private Property Let IGetLet_Value(ByVal    NewValue      As Variant) : Let Value = NewValue           : End Property
Private Property Let IGetLetSet_Value(ByVal NewValue      As Variant) : Let Value = NewValue           : End Property
Private Property Set ISet_Value(ByVal       NewValue      As Variant) : Set Value = NewValue           : End Property
Private Property Set IGetSet_Value(ByVal    NewValue      As Variant) : Set Value = NewValue           : End Property
Private Property Set IGetLetSet_Value(ByVal NewValue      As Variant) : Set Value = NewValue           : End Property
Private Property Get IGet_Value()                         As Variant  : Call Extract(IGet_Value)       : End Property
Private Property Get IGetLet_Value()                      As Variant  : Call Extract(IGetLet_Value)    : End Property
Private Property Get IGetSet_Value()                      As Variant  : Call Extract(IGetSet_Value)    : End Property
Private Property Get IGetLetSet_Value()                   As Variant  : Call Extract(IGetLetSet_Value) : End Property
Private Sub          IGetLetSet_Assign(ByVal Value        As Variant) : Call Assign(Value)             : End Sub
Private Sub          IGetLetSet_Extract(ByRef Destination As Variant) : Call Extract(Destination)      : End Sub
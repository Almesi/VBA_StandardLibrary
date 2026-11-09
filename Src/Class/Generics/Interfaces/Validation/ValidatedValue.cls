VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ValidatedValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mValue    As Variant
Private mDataType As CoreValueType

Public Function Create(ByVal Value As Variant, ByVal DataType As CoreValueType) As ValidatedValue
    Set Create = New ValidatedValue
    Call Create.Initialize(Value, DataType)
End Function

Friend Sub Initialize(ByVal Value As Variant, ByVal DataType As CoreValueType)
    If IsObject(Value) Then
        Set mValue = Value
    Else
        Let mValue = Value
    End If
    mDataType = DataType
End Sub

Public Property Get Value() As Variant
    If IsObject(Value) Then
        Set Value = mValue
    Else
        Let Value = mValue
    End If
End Property

Public Property Get DataType() As CoreValueType
    DataType = mDataType
End Property
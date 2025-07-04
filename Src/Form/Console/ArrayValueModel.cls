VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayValueModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit


Private p_Arr() As Variant

Public Property Let Arr(n_Arr() As Variant)
    p_Arr = n_Arr
End Property
Public Property Get Arr() As Variant()
    Arr = p_Arr
End Property

Public Function Count() As Long
    Count = ArraySize(Arr)
End Function

Public Property Let Element(Index As Long, n_Value As Variant)
    If IsObject(n_Value) Then
        Set p_Arr(Index) = n_Value
    Else
        p_Arr(Index) = n_Value
    End If
End Property
Public Property Get Element(Index As Long) As Variant
    If IsObject(p_Arr(Index)) Then
        Set Element = p_Arr(Index)
    Else
        Element = p_Arr(Index)
    End If
End Property

Public Function Add(n_Value As Variant) As Long
    ReDim Preserve p_Arr(ArraySize(p_Arr) + 1)
    Element(ArraySize(p_Arr)) = n_Value
    Add = ArraySize(p_Arr)
End Function

Private Function ArraySize(Arr As Variant, Optional Dimension As Long = 1) As Long
    On Error Resume Next
    ArraySize = -1
    ArraySize = Ubound(Arr, Dimension)
End Function
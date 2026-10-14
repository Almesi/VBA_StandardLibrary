VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetValueDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IGetValue

Private mValue As Double

Public Function Create(ByVal Value As Double) As GetValueDouble
    mValue = Value
    Set Create = Me
End Function

Private Function IGetValue_GetValue() As Variant
    GetValue = mValue
End Function

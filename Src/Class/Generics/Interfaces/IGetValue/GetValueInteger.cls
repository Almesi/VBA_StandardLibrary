VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetValueInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IGetValue

Private mValue As Integer

Public Function Create(ByVal Value As Integer) As GetValueInteger
    mValue = Value
    Set Create = Me
End Function

Private Function IGetValue_GetValue() As Variant
    GetValue = mValue
End Function

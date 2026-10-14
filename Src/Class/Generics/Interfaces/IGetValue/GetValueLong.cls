VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetValueLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IGetValue

Private mValue As Long

Public Function Create(ByVal Value As Long) As GetValueLong
    mValue = Value
    Set Create = Me
End Function

Private Function IGetValue_GetValue() As Variant
    GetValue = mValue
End Function

Attribute VB_Name = "std_GeneralFunctions"


Option Explicit

Public Sub Assign(ByRef Destination As Variant, ByVal Value As Variant)
    If IsObject(Value) Then
        Set Destination = Value
    Else
        Let Destination = Value
    End If
End Sub

Public Function USize(ByRef Arr As Variant) As Long
    On Error Resume Next
    USize = -1
    USize = Ubound(Arr)
End Function
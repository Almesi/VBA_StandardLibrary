VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "RangeModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TModel
    Range As Range
End Type

Private This As TModel


Public Property Get Range() As Range
    Set Range = This.Range
End Property

Public Property Set Range(ByVal Value As Range)
    Set This.Range = Value
End Property


Public Property Get Address() As String
    Address = This.Range.Address
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = This.Range.Worksheet
End Property


Public Property Get Self() As RangeModel
    Set Self = Me
End Property
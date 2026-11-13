VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGetLetSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public Property Get Value() As Variant
Public Property Let Value(ByVal NewValue As Variant)
Public Property Set Value(ByVal NewValue As Variant)
Public Sub          Assign(ByVal Value As Variant)
Public Sub          Extract(ByRef Destination As Variant)
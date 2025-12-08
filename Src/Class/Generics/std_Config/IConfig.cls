VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Public Property Let Destination(n_Destination As Variant)
End Property

Public Property Get Destination() As Variant
End Property

Public Property Let Setting(Indicator As Variant, Optional Check As Boolean = False, Optional Index As Long = 0, Value As Variant)
End Property

Public Property Get Setting(Indicator As Variant, Optional Check As Boolean = False, Optional Index As Long = 0) As Variant
End Property

Public Function Check(Scope As Variant, Value As Variant) As Boolean
End Function
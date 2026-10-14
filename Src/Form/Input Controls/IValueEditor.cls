VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public Property Get Control()             As MSForms.Control
Public Property Let Value(ByVal Value     As Variant)
Public Property Get Value()               As Variant
Public Function     Validate()            As Boolean
Public Function     Convert()             As Variant
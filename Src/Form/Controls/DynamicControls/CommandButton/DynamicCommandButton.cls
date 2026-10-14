VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicCommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IControlValue

Private mControl           As MSForms.CommandButton
Private mEvents            As CommandButtonEvents

Public Function Create(ByVal nControl As MSForms.CommandButton, ByVal nEvents As CommandButtonEvents) As DynamicCommandButton
    Set Create = New DynamicCommandButton
    Call Create.Initialize(nControl, nEvents)
End Function

Friend Sub Initialize(ByVal nControl As MSForms.CommandButton, ByVal nEvents As CommandButtonEvents)
    Set mControl = nControl
    Set mEvents  = nEvents
End Sub

Public Property Get Control() As MSForms.CommandButton
    Set Control = mControl
End Property

Public Property Get Events() As CommandButtonEvents
    Set Events = mEvents
End Property



Private Property Let IControlValue_Value(ByVal nValue As Variant)
    mControl.Value = nValue
End Property

Private Property Get IControlValue_Value() As Variant
    IControlValue_Value = mControl.Value
End Function
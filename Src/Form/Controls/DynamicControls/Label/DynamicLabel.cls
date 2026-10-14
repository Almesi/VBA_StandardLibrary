VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IControlValue

Private mControl           As MSForms.Label
Private mEvents            As LabelEvents

Public Function Create(ByVal nControl As MSForms.Label, ByVal nEvents As LabelEvents) As DynamicLabel
    Set Create = New DynamicLabel
    Call Create.Initialize(nControl, nEvents)
End Function

Friend Sub Initialize(ByVal nControl As MSForms.Label, ByVal nEvents As LabelEvents)
    Set mControl = nControl
    Set mEvents  = nEvents
End Sub

Public Property Get Control() As MSForms.Label
    Set Control = mControl
End Property

Public Property Get Events() As LabelEvents
    Set Events = mEvents
End Property



Private Property Let IControlValue_Value(ByVal nValue As Variant)
    mControl.Value = nValue
End Property

Private Property Get IControlValue_Value() As Variant
    IControlValue_Value = mControl.Value
End Function
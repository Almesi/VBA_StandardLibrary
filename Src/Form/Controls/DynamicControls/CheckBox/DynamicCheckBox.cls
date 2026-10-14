VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IControlValue

Private mControl           As MSForms.CheckBox
Private mEvents            As CheckBoxEvents

Public Function Create(ByVal nControl As MSForms.CheckBox, ByVal nEvents As CheckBoxEvents) As DynamicCheckBox
    Set Create = New DynamicCheckBox
    Call Create.Initialize(nControl, nEvents)
End Function

Friend Sub Initialize(ByVal nControl As MSForms.CheckBox, ByVal nEvents As CheckBoxEvents)
    Set mControl = nControl
    Set mEvents  = nEvents
End Sub

Public Property Get Control() As MSForms.CheckBox
    Set Control = mControl
End Property

Public Property Get Events() As CheckBoxEvents
    Set Events = mEvents
End Property



Private Property Let IControlValue_Value(ByVal nValue As Variant)
    mControl.Value = nValue
End Property

Private Property Get IControlValue_Value() As Variant
    IControlValue_Value = mControl.Value
End Function
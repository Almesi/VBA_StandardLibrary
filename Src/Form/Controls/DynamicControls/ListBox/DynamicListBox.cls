VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IControlValue

Private mControl           As MSForms.ListBox
Private mEvents            As ListBoxEvents

Public Function Create(ByVal nControl As MSForms.ListBox, ByVal nEvents As ListBoxEvents) As DynamicListBox
    Set Create = New DynamicListBox
    Call Create.Initialize(nControl, nEvents)
End Function

Friend Sub Initialize(ByVal nControl As MSForms.ListBox, ByVal nEvents As ListBoxEvents)
    Set mControl = nControl
    Set mEvents  = nEvents
End Sub

Public Property Get Control() As MSForms.ListBox
    Set Control = mControl
End Property

Public Property Get Events() As ListBoxEvents
    Set Events = mEvents
End Property



Private Property Let IControlValue_Value(ByVal nValue As Variant)
    mControl.Value = nValue
End Property

Private Property Get IControlValue_Value() As Variant
    IControlValue_Value = mControl.Value
End Function
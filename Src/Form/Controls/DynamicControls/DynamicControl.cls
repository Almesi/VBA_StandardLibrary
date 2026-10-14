VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pControl  As Object
Private pEvents   As Object

Public Function Create(ByVal Control As Object, ByVal Events As Object) As DynamicControl
    Set Create = New DynamicControl
    Call Create.Initialize(Control, Events)
End Function

Public Function CreateFromValid(ByVal Valid As Object) As DynamicControl
    Set CreateFromValid = New DynamicControl
    Call CreateFromValid.Initialize(Valid.Control, Valid.Events)
End Function

Friend Sub Initialize(ByVal Control As Object, ByVal Events As Object)
    Set pControl = Control
    Set pEvents  = Events
End Sub

Public Property Get Control() As Object
    Set Control = pControl
End Property

Public Property Get Events() As Object
    Set Events = pEvents
End Property
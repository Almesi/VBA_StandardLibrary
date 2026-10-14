VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicControlFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pFactory As Object

Public Function Create(ByVal Factory As Object) As DynamicControlFactory
    Set Create = New DynamicControlFactory
    Call Create.Initialize(Factory)
End Function

Public Function CreateFromValid(ByVal Valid As Object) As DynamicControlFactory
    Set CreateFromValid = New DynamicControlFactory
    Call CreateFromValid.Initialize(Valid)
End Function

Friend Sub Initialize(ByVal Factory As Object)
    Set pFactory = Factory
End Sub

Public Function CreateControl(ByVal Destination As Object) As DynamicControl
    Dim Valid As Object
    Set Valid = pFactory.CreateControl(Destination)
    Set CreateControl = DynamicControl.CreateFromValid(Valid)
End Function
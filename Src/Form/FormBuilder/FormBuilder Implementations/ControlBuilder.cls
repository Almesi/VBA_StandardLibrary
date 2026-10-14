VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "ControlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IFormBuilder

Private mFactory As DynamicControlFactory

Public Function Create(ByVal Factory As DynamicControlFactory) As ControlBuilder
    Set Create = New ControlBuilder
    Call Create.Initialize(Factory)
End Function


Friend Sub Initialize(ByVal Factory As DynamicControlFactory)
    Set mFactory = Factory
End Sub

Private Sub IFormBuilder_Build(ByVal Context As FormBuilderContext, ByVal Scope As FormBuilderScope)
    Dim LifeTimeIndex As Long
    LifeTimeIndex = Context.Form.AddLifeTime(mFactory.CreateControl(Context.Form.Form))
    Call Scope.AddNumber(LifeTimeIndex)
End Sub
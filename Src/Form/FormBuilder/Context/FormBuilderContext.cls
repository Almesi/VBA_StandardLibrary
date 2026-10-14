VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormBuilderContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mForm  As FormBuilderForm
Private mScope As FormBuilderScope


Public Function Create(ByVal Form As FormBuilderForm) As FormBuilderContext
    Set Create = New FormBuilderContext
    Call Create.Initialize(Form)
End Function

Friend Sub Initialize(ByVal Form As FormBuilderForm)
    Set mForm = Form
    Set mScope = FormBuilderScope.Create(Form)
End Sub

Public Property Get Form() As FormBuilderForm
    Set Form = mForm
End Property

Public Property Get Scope() As FormBuilderScope
    Set Scope = mScope
End Property
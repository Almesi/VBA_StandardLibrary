VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormBuilderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mForm     As Object
Private mLifeTime As Collection

Public Function Create(ByVal Form As Object) As FormBuilderForm
    Set Create = New FormBuilderForm
    Call Create.Initialize(Form)
End Function

Friend Sub Initialize(ByVal Form As Object)
    Set mForm     = Form
    Set mLifeTime = New Collection
End Sub

Public Property Get Form() As Object
    Set Form = mForm
End Property

Public Property Get LifeTime() As Collection
    Set LifeTime = mLifeTime
End Property

Public Function AddLifeTime(ByVal Value As Object) As Long
    Call mLifeTime.Add(Value)
    AddLifeTime = mLifeTime.Count
End Function
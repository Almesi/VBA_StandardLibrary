VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormBuilderScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mForm   As FormBuilderForm
Private mParent As FormBuilderScope
Private mNodes  As Collection


Public Function Create(ByVal Form As FormBuilderForm, Optional ByVal Parent As FormBuilderScope) As FormBuilderScope
    Set Create = New FormBuilderScope
    Call Create.Initialize(Form, Parent)
End Function


Friend Sub Initialize(ByVal Form As FormBuilderForm, Optional ByVal Parent As FormBuilderScope)
    Set mForm   = Form
    Set mParent = Parent
    Set mNodes  = New Collection
End Sub


Public Sub AddNumber(ByVal LifeTimeIndex As Long)
    Call mNodes.Add(LifeTimeIndex)
End Sub


Public Sub AddScope(ByVal Scope As FormBuilderScope)
    Call mNodes.Add(Scope)
End Sub


Public Property Get Parent() As FormBuilderScope
    Set Parent = mParent
End Property


Public Property Get Count() As Long
    Count = mNodes.Count
End Property


Public Property Get Node(ByVal Position As Long) As Variant
    If IsObject(mNodes.Item(Position)) Then
        Set Node = mNodes.Item(Position)
    Else
        Node = mNodes.Item(Position)
    End If
End Property
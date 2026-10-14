VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PositionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IFormBuilder

Private mLeft As Single
Private mTop As Single
Private mBuilder As IFormBuilder


Public Function Create(ByVal Builder As IFormBuilder, ByVal Left As Single, ByVal Top As Single) As PositionBuilder
    Set Create = New PositionBuilder
    Call Create.Initialize(Builder, Left, Top)
End Function


Friend Sub Initialize(ByVal Builder As IFormBuilder, ByVal Left As Single, ByVal Top As Single)
    mLeft = Left
    mTop  = Top
    Set mBuilder = Builder
End Sub


Private Sub IFormBuilder_Build(ByVal Context As FormBuilderContext, ByVal Scope As FormBuilderScope)
    Call mBuilder.Build(Context, Scope)
    Call ApplyPosition(Scope)
End Sub

Private Sub ApplyPosition(ByVal Scope As FormBuilderScope)
    Dim i       As Long
    Dim Control As Object

    For i = 1 To Scope.Count
        If IsObject(Scope.Node(i)) Then
            Call ApplyPosition(Scope.Node(i))
        Else
            Set Control  = Scope.Node(i).Control
            Control.Left = Control.Left + mLeft
            Control.Top  = Control.Top  + mTop
        End If
    Next i
End Sub
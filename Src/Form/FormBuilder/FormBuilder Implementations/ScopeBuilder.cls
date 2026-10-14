VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "ScopeBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IFormBuilder

Private mBuilder As IFormBuilder


Public Function Create(ByVal Builder As IFormBuilder) As ScopeBuilder
    Set Create = New ScopeBuilder
    Call Create.Initialize(Builder)
End Function

Friend Sub Initialize(ByVal Builder As IFormBuilder)
    Set mBuilder = Builder
End Sub

Private Sub IFormBuilder_Build(ByVal Context As FormBuilderContext, ByVal Scope As FormBuilderScope)
    Dim ChildScope As FormBuilderScope
    Set ChildScope = FormBuilderScope.Create(Context.Form, Scope)
    Call Scope.AddScope(ChildScope)
    Call mBuilder.Build(Context, ChildScope)
End Sub
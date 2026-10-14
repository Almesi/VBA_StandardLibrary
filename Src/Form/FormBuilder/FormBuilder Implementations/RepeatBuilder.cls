VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "RepeatBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IFormBuilder

Private mBuilder As IFormBuilder
Private mCount As Long

Public Function Create(ByVal Builder As IFormBuilder, ByVal Count As Long) As RepeatBuilder
    Set Create = New RepeatBuilder
    Call Create.Initialize(Builder, Count)
End Function


Friend Sub Initialize(ByVal Builder As IFormBuilder, ByVal Count As Long)
    Set mBuilder = Builder
    mCount = Count
End Sub

Private Sub IFormBuilder_Build(ByVal Context As FormBuilderContext, ByVal Scope As FormBuilderScope)
    Dim i As Long

    For i = 1 To mCount
        Call mBuilder.Build(Context, Scope)
    Next i
End Sub
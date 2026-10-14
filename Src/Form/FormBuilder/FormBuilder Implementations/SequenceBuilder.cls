VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "SequenceBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IFormBuilder

Private mBuilders As Collection

Public Function Create() As SequenceBuilder
    Set Create = New SequenceBuilder
    Call Create.Initialize
End Function


Friend Sub Initialize()
    Set mBuilders = New Collection
End Sub


Public Sub Add(ByVal Builder As IFormBuilder)
    Call mBuilders.Add(Builder)
End Sub


Public Property Get Count() As Long
    Count = mBuilders.Count
End Property


Public Property Get Item(ByVal Position As Long) As IFormBuilder
    Set Item = mBuilders.Item(Position)
End Property


Private Sub IFormBuilder_Build(ByVal Context As FormBuilderContext, ByVal Scope As FormBuilderScope)
    Dim Builder As IFormBuilder

    For Each Builder In mBuilders
        Call Builder.Build(Context, Scope)
    Next Builder
End Sub
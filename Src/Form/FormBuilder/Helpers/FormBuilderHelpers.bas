Attribute VB_Name = "FormBuilderHelpers"


Option Explicit

' ============================================================================
' IFormBuilder Composition Helpers
' ============================================================================

Public Function FormBuilderControl(ByVal Factory As Object) As IFormBuilder
    Set FormBuilderControl = ControlBuilder.Create(DynamicControlFactory.Create(Factory))
End Function


Public Function FormBuilderSequence(ParamArray Builders() As Variant) As IFormBuilder
    Dim Result As SequenceBuilder
    Dim i      As Long

    Set Result = SequenceBuilder.Create

    For i = 0 To UBound(Builders)
        Call Result.Add(Builders(i))
    Next i
    Set FormBuilderSequence = Result
End Function


Public Function FormBuilderColumn(ByVal Builder As IFormBuilder, Optional ByVal PadWidth As Single = 0, Optional ByVal PadHeight As Single = 0) As IFormBuilder
    Set FormBuilderColumn = LayoutBuilder.Create(Builder,LayoutColumn,PadWidth,PadHeight)
End Function


Public Function FormBuilderRow(ByVal Builder As IFormBuilder, Optional ByVal PadWidth As Single = 0, Optional ByVal PadHeight As Single = 0) As IFormBuilder
    Set FormBuilderRow = LayoutBuilder.Create(Builder,LayoutRow,PadWidth,PadHeight)
End Function


Public Function FormBuilderScope(ByVal Builder As IFormBuilder) As IFormBuilder
    Set FormBuilderScope = ScopeBuilder.Create(Builder)
End Function


Public Function FormBuilderPosition(ByVal Builder As IFormBuilder,ByVal Left As Single,ByVal Top As Single) As IFormBuilder
    Set FormBuilderPosition = PositionBuilder.Create(Builder,Left,Top)
End Function


Public Function FormBuilderRepeat(ByVal Builder As IFormBuilder,ByVal Count As Long) As IFormBuilder
    Set FormBuilderRepeat = RepeatBuilder.Create(Builder,Count)
End Function

Public Function FormBuilderControls( _
    ParamArray Factories() As Variant _
) As IFormBuilder

    Dim Temp() As Variant
    Temp = Factories
    Set FormBuilderControls = FormBuilderControlsArg(Temp)

End Function

Public Function FormBuilderControlsArg( _
    ByVal Factories As Variant _
) As IFormBuilder

    Dim Builder As SequenceBuilder
    Dim i       As Long

    Set Builder = SequenceBuilder.Create

    For i = 0 To UBound(Factories)
        Builder.Add FormBuilderControl(Factories(i))
    Next i

    Set FormBuilderControlsArg = Builder

End Function
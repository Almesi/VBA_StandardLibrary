Attribute VB_Name = "FormBuilderPreBuild1"

Option Explicit

Public Function LabelWithRow( _
             ByVal Content As IFormBuilder, _
    Optional ByVal LabelFactory As Object = Nothing, _
    Optional ByVal PadWidth As Single = 0 _
) As IFormBuilder

    Dim LabelBuilder As IFormBuilder

    If LabelFactory Is Nothing Then
        Set LabelBuilder = FormBuilderControl(CreateLabelFactory(""))
    Else
        Set LabelBuilder = FormBuilderControl(LabelFactory)
    End If

    Set LabelWithRow = FormBuilderRow( _
        FormBuilderSequence( _
            LabelBuilder, _
            Content _
        ), _
        PadWidth _
    )

End Function

Public Function LabelWithColumn( _
             ByVal Content As IFormBuilder, _
    Optional ByVal LabelFactory As Object = Nothing, _
    Optional ByVal PadHeight As Single = 0 _
) As IFormBuilder

    Dim LabelBuilder As IFormBuilder

    If LabelFactory Is Nothing Then
        Set LabelBuilder = FormBuilderControl(CreateLabelFactory(""))
    Else
        Set LabelBuilder = FormBuilderControl(LabelFactory)
    End If

    Set LabelWithColumn = FormBuilderColumn( _
        FormBuilderSequence( _
            LabelBuilder, _
            Content _
        ), _
        0, _
        PadHeight _
    )

End Function

'=========================================================================
' Labelled controls
'=========================================================================

Public Function LabelWithRowControls( _
             ByVal LabelFactory As Object, _
             ParamArray ControlFactories() As Variant _
) As IFormBuilder

    Dim Temp As Variant
    Temp = ControlFactories

    Set LabelWithRowControls = LabelWithRowControlsArg( _
        LabelFactory, _
        Temp _
    )

End Function

Public Function LabelWithRowControlsArg( _
             ByVal LabelFactory As Object, _
             ByVal ControlFactories As Variant _
) As IFormBuilder

    Set LabelWithRowControlsArg = LabelWithRow( _
        FormBuilderControlsArg(ControlFactories), _
        LabelFactory _
    )

End Function

Public Function LabelWithColumnControls( _
             ByVal LabelFactory As Object, _
             ParamArray ControlFactories() As Variant _
) As IFormBuilder

    Dim Temp As Variant
    Temp = ControlFactories

    Set LabelWithColumnControls = LabelWithColumnControlsArg( _
        LabelFactory, _
        Temp _
    )

End Function

Public Function LabelWithColumnControlsArg( _
             ByVal LabelFactory As Object, _
             ByVal ControlFactories As Variant _
) As IFormBuilder

    Set LabelWithColumnControlsArg = LabelWithColumn( _
        FormBuilderColumn( _
            FormBuilderControlsArg(ControlFactories), _
            0, _
            5 _
        ), _
        LabelFactory _
    )

End Function

Public Function ButtonPair( _
             ByVal FirstFactory As Object, _
             ByVal SecondFactory As Object, _
    Optional ByVal PadWidth As Single = 0 _
) As IFormBuilder

    Set ButtonPair = FormBuilderRow( _
        FormBuilderControls( _
            FirstFactory, _
            SecondFactory _
        ), _
        PadWidth _
    )

End Function

Public Function OKCancelButtons( _
             ByVal OKFactory As Object, _
             ByVal CancelFactory As Object, _
    Optional ByVal PadWidth As Single = 0 _
) As IFormBuilder

    Set OKCancelButtons = ButtonPair( _
        OKFactory, _
        CancelFactory, _
        PadWidth _
    )

End Function

Public Function YesNoButtons( _
             ByVal YesFactory As Object, _
             ByVal NoFactory As Object, _
    Optional ByVal PadWidth As Single = 0 _
) As IFormBuilder

    Set YesNoButtons = ButtonPair( _
        YesFactory, _
        NoFactory, _
        PadWidth _
    )

End Function

Public Function BackNextButtons( _
             ByVal BackFactory As Object, _
             ByVal NextFactory As Object, _
    Optional ByVal PadWidth As Single = 0 _
) As IFormBuilder

    Set BackNextButtons = ButtonPair( _
        BackFactory, _
        NextFactory, _
        PadWidth _
    )

End Function

Public Function FieldControl( _
             ByVal Content As IFormBuilder, _
    Optional ByVal LabelFactory As Object = Nothing, _
    Optional ByVal PadHeight As Single = 0 _
) As IFormBuilder

    Dim LabelBuilder As IFormBuilder

    If LabelFactory Is Nothing Then
        Set LabelBuilder = FormBuilderControl(CreateLabelFactory(""))
    Else
        Set LabelBuilder = FormBuilderControl(LabelFactory)
    End If

    Set FieldControl = FormBuilderColumn( _
        FormBuilderSequence( _
            LabelBuilder, _
            Content _
        ), _
        0, _
        PadHeight _
    )

End Function

Public Function ControlRow( _
             ByVal Factory As Object, _
             ByVal Count As Long, _
    Optional ByVal PadWidth As Single = 0 _
) As IFormBuilder

    Set ControlRow = FormBuilderRow( _
        FormBuilderRepeat( _
            FormBuilderControl(Factory), _
            Count _
        ), _
        PadWidth _
    )

End Function

Public Function ControlColumn( _
             ByVal Factory As Object, _
             ByVal Count As Long, _
    Optional ByVal PadHeight As Single = 0 _
) As IFormBuilder

    Set ControlColumn = FormBuilderColumn( _
        FormBuilderRepeat( _
            FormBuilderControl(Factory), _
            Count _
        ), _
        0, _
        PadHeight _
    )

End Function

'=========================================================================
' Multiple controls
'=========================================================================

Public Function ControlsRow( _
             ParamArray ControlFactories() As Variant _
) As IFormBuilder

    Dim Temp As Variant
    Temp = ControlFactories

    Set ControlsRow = ControlsRowArg(Temp)

End Function

Public Function ControlsRowArg( _
             ByVal ControlFactories As Variant _
) As IFormBuilder

    Set ControlsRowArg = FormBuilderRow( _
        FormBuilderControlsArg(ControlFactories), _
        5 _
    )

End Function

Public Function ControlsColumn( _
             ParamArray ControlFactories() As Variant _
) As IFormBuilder

    Dim Temp As Variant
    Temp = ControlFactories

    Set ControlsColumn = ControlsColumnArg(Temp)

End Function

Public Function ControlsColumnArg( _
             ByVal ControlFactories As Variant _
) As IFormBuilder

    Set ControlsColumnArg = FormBuilderColumn( _
        FormBuilderControlsArg(ControlFactories), _
        0, _
        5 _
    )

End Function

Public Function LabeledButtonGroup( _
             ByVal LabelFactory As Object, _
             ParamArray ButtonFactories() As Variant _
) As IFormBuilder

    Dim Temp As Variant
    Temp = ButtonFactories

    Set LabeledButtonGroup = LabeledButtonGroupArg( _
        LabelFactory, _
        Temp _
    )

End Function

Public Function LabeledButtonGroupArg( _
             ByVal LabelFactory As Object, _
             ByVal ButtonFactories As Variant _
) As IFormBuilder

    Set LabeledButtonGroupArg = LabelWithRow( _
        FormBuilderControlsArg(ButtonFactories), _
        LabelFactory _
    )

End Function

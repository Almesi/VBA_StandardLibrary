Attribute VB_Name = "FactoryHelpers"


Option Explicit

' ============================================================================
' Control Typee postfixes
' ============================================================================

Public Enum ControlType
    ControlTypeLabel
    ControlTypeTextBox
    ControlTypeCommandButton
    ControlTypeComboBox
    ControlTypeCheckBox
    ControlTypeOptionButton
    ControlTypeListBox
    ControlTypeFrame
    ControlTypeImage
    ControlTypeScrollBar
    ControlTypeSpinButton
    ControlTypeToggleButton
End Enum


' ============================================================================
' Standard control dimensions
'
' Change these values to change the standard dimensions throughout the
' abstraction layer.
' ============================================================================

Public Const STANDARD_LABEL_WIDTH         As Single = 100
Public Const STANDARD_LABEL_HEIGHT        As Single = 30
Public Const STANDARD_TEXTBOX_WIDTH       As Single = 100
Public Const STANDARD_TEXTBOX_HEIGHT      As Single = 30
Public Const STANDARD_BUTTON_WIDTH        As Single = 50
Public Const STANDARD_BUTTON_HEIGHT       As Single = 30
Public Const STANDARD_COMBOBOX_WIDTH      As Single = 100
Public Const STANDARD_COMBOBOX_HEIGHT     As Single = 30
Public Const STANDARD_CHECKBOX_WIDTH      As Single = 100
Public Const STANDARD_CHECKBOX_HEIGHT     As Single = 30
Public Const STANDARD_OPTIONBUTTON_WIDTH  As Single = 100
Public Const STANDARD_OPTIONBUTTON_HEIGHT As Single = 30
Public Const STANDARD_LISTBOX_WIDTH       As Single = 100
Public Const STANDARD_LISTBOX_HEIGHT      As Single = 100
Public Const STANDARD_FRAME_WIDTH         As Single = 100
Public Const STANDARD_FRAME_HEIGHT        As Single = 100


' ============================================================================
' Standard spacing
' ============================================================================

Public Const STANDARD_HORIZONTAL_PADDING As Single = 10
Public Const STANDARD_VERTICAL_PADDING   As Single = 10
Public Const STANDARD_LABEL_CONTROL_GAP  As Single = 0
Public Const STANDARD_BUTTON_GAP         As Single = 0


' ============================================================================
' Naming
' ============================================================================

Public Function ControlName(ByVal BaseName As String, ByVal Typee As ControlType) As String
    ControlName = BaseName & ControlTypePostfix(Typee)
End Function


Private Function ControlTypePostfix(ByVal Typee As ControlType) As String
    Select Case Typee
        Case ControlTypeLabel         : ControlTypePostfix = "Label"
        Case ControlTypeTextBox       : ControlTypePostfix = "TextBox"
        Case ControlTypeCommandButton : ControlTypePostfix = "Button"
        Case ControlTypeComboBox      : ControlTypePostfix = "ComboBox"
        Case ControlTypeCheckBox      : ControlTypePostfix = "CheckBox"
        Case ControlTypeOptionButton  : ControlTypePostfix = "OptionButton"
        Case ControlTypeListBox       : ControlTypePostfix = "ListBox"
        Case ControlTypeFrame         : ControlTypePostfix = "Frame"
        Case ControlTypeImage         : ControlTypePostfix = "Image"
        Case ControlTypeScrollBar     : ControlTypePostfix = "ScrollBar"
        Case ControlTypeSpinButton    : ControlTypePostfix = "SpinButton"
        Case ControlTypeToggleButton  : ControlTypePostfix = "ToggleButton"
        Case Else
            Err.Raise _
                vbObjectError + 1000, _
                "FormBuilderHelpers.ControlTypePostfix", _
                "Unknown control Typee."
    End Select
End Function


' ============================================================================
' Common factory configuration
' ============================================================================

Public Sub ApplyName(ByVal Factory As Object, ByVal BaseName As String, ByVal Typee As ControlType)
    Factory.Common.Name = ControlName(BaseName, Typee)
End Sub


Public Sub ApplyPosition(ByVal Factory As Object, ByVal Left As Single, ByVal Top As Single)
    Factory.Common.Left = Left
    Factory.Common.Top  = Top
End Sub


Public Sub ApplySize(ByVal Factory As Object, ByVal Width As Single, ByVal Height As Single)
    Factory.Common.Width  = Width
    Factory.Common.Height = Height
End Sub


Public Sub ApplyNamePosition( _
    ByVal Factory As Object, _
    ByVal BaseName As String, _
    ByVal Typee As ControlType, _
    ByVal Left As Single, _
    ByVal Top As Single _
)

    Call ApplyName(Factory, BaseName, Typee)
    Call ApplyPosition(Factory, Left, Top)
End Sub


Public Sub ApplyNamePositionSize( _
    ByVal Factory As Object, _
    ByVal BaseName As String, _
    ByVal Typee As ControlType, _
    ByVal Left As Single, _
    ByVal Top As Single, _
    ByVal Width As Single, _
    ByVal Height As Single _
)

    Call ApplyName(Factory, BaseName, Typee)
    Call ApplyPosition(Factory, Left, Top)
    Call ApplySize(Factory, Width, Height)

End Sub


' ============================================================================
' Standard positioning
'
' Applies the standard size associated with the control Typee as well as
' its name and position.
' ============================================================================

Public Sub ApplyStandardPosition( _
    ByVal Factory As Object, _
    ByVal BaseName As String, _
    ByVal Typee As ControlType, _
    ByVal Left As Single, _
    ByVal Top As Single _
)

    Call ApplyName(Factory, BaseName, Typee)
    Call ApplyPosition(Factory, Left, Top)
    Call ApplyStandardSize(Factory, Typee)
End Sub


Public Sub ApplyStandardSize(ByVal Factory As Object, ByVal Typee As ControlType)
    Select Case Typee
        Case ControlTypeLabel         : Call ApplySize(Factory, STANDARD_LABEL_WIDTH        , STANDARD_LABEL_HEIGHT)
        Case ControlTypeTextBox       : Call ApplySize(Factory, STANDARD_TEXTBOX_WIDTH      , STANDARD_TEXTBOX_HEIGHT)
        Case ControlTypeCommandButton : Call ApplySize(Factory, STANDARD_BUTTON_WIDTH       , STANDARD_BUTTON_HEIGHT)
        Case ControlTypeComboBox      : Call ApplySize(Factory, STANDARD_COMBOBOX_WIDTH     , STANDARD_COMBOBOX_HEIGHT)
        Case ControlTypeCheckBox      : Call ApplySize(Factory, STANDARD_CHECKBOX_WIDTH     , STANDARD_CHECKBOX_HEIGHT)
        Case ControlTypeOptionButton  : Call ApplySize(Factory, STANDARD_OPTIONBUTTON_WIDTH , STANDARD_OPTIONBUTTON_HEIGHT)
        Case ControlTypeListBox       : Call ApplySize(Factory, STANDARD_LISTBOX_WIDTH      , STANDARD_LISTBOX_HEIGHT)
        Case ControlTypeFrame         : Call ApplySize(Factory, STANDARD_FRAME_WIDTH        , STANDARD_FRAME_HEIGHT)
        Case Else
    End Select
End Sub


' ============================================================================
' Complete standard configuration
' ============================================================================

Public Sub ApplyStandard( _
    ByVal Factory As Object, _
    ByVal BaseName As String, _
    ByVal Typee As ControlType, _
    ByVal Left As Single, _
    ByVal Top As Single _
)

    Call ApplyStandardPosition(Factory, BaseName, Typee, Left, Top)
End Sub



Public Function CreateLabelFactory( _
    ByVal Name As String, _
    Optional ByVal Caption As String = "Label" _
) As Object

    Dim Factory As DynamicLabelFactory

    Set Factory = DynamicLabelFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeLabel, _
        0, _
        0 _
    )

    Factory.Caption = Caption

    Set CreateLabelFactory = Factory

End Function


Public Function CreateTextBoxFactory( _
    ByVal Name As String _
) As Object

    Dim Factory As DynamicTextBoxFactory

    Set Factory = DynamicTextBoxFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeTextBox, _
        0, _
        0 _
    )

    Set CreateTextBoxFactory = Factory

End Function


Public Function CreateButtonFactory( _
    ByVal Name As String, _
    Optional ByVal Caption As String = "Button" _
) As Object

    Dim Factory As DynamicCommandButtonFactory

    Set Factory = DynamicCommandButtonFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeCommandButton, _
        0, _
        0 _
    )

    Factory.Caption = Caption

    Set CreateButtonFactory = Factory

End Function


Public Function CreateComboBoxFactory( _
    ByVal Name As String _
) As Object

    Dim Factory As DynamicComboBoxFactory

    Set Factory = DynamicComboBoxFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeComboBox, _
        0, _
        0 _
    )

    Set CreateComboBoxFactory = Factory

End Function


Public Function CreateCheckBoxFactory( _
    ByVal Name As String, _
    Optional ByVal Caption As String = "CheckBox" _
) As Object

    Dim Factory As DynamicCheckBoxFactory

    Set Factory = DynamicCheckBoxFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeCheckBox, _
        0, _
        0 _
    )

    Factory.Caption = Caption

    Set CreateCheckBoxFactory = Factory

End Function


Public Function CreateOptionButtonFactory( _
    ByVal Name As String, _
    Optional ByVal Caption As String = "Option" _
) As Object

    Dim Factory As DynamicOptionButtonFactory

    Set Factory = DynamicOptionButtonFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeOptionButton, _
        0, _
        0 _
    )

    Factory.Caption = Caption

    Set CreateOptionButtonFactory = Factory

End Function


Public Function CreateListBoxFactory( _
    ByVal Name As String _
) As Object

    Dim Factory As DynamicListBoxFactory

    Set Factory = DynamicListBoxFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeListBox, _
        0, _
        0 _
    )

    Set CreateListBoxFactory = Factory

End Function


Public Function CreateFrameFactory( _
    ByVal Name As String, _
    Optional ByVal Caption As String = "Frame" _
) As Object

    Dim Factory As DynamicFrameFactory

    Set Factory = DynamicFrameFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeFrame, _
        0, _
        0 _
    )

    Factory.Caption = Caption

    Set CreateFrameFactory = Factory

End Function


Public Function CreateImageFactory( _
    ByVal Name As String _
) As Object

    Dim Factory As DynamicImageFactory

    Set Factory = DynamicImageFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeImage, _
        0, _
        0 _
    )

    Set CreateImageFactory = Factory

End Function


Public Function CreateScrollBarFactory( _
    ByVal Name As String _
) As Object

    Dim Factory As DynamicScrollBarFactory

    Set Factory = DynamicScrollBarFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeScrollBar, _
        0, _
        0 _
    )

    Set CreateScrollBarFactory = Factory

End Function


Public Function CreateSpinButtonFactory( _
    ByVal Name As String _
) As Object

    Dim Factory As DynamicSpinButtonFactory

    Set Factory = DynamicSpinButtonFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeSpinButton, _
        0, _
        0 _
    )

    Set CreateSpinButtonFactory = Factory

End Function


Public Function CreateToggleButtonFactory( _
    ByVal Name As String, _
    Optional ByVal Caption As String = "" _
) As Object

    Dim Factory As DynamicToggleButtonFactory

    Set Factory = DynamicToggleButtonFactory.Create

    Call ApplyStandard( _
        Factory, _
        Name, _
        ControlTypeToggleButton, _
        0, _
        0 _
    )

    Factory.Caption = Caption

    Set CreateToggleButtonFactory = Factory

End Function
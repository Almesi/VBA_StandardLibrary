VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "EnumValueEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IValueEditor

Private ControlValue As MSForms.ComboBox
Private Definition   As FieldDefinition
Private IsEditable   As Boolean

Public Function Create(ByVal Parent As Object, ByVal FieldDefinition As FieldDefinition, ByVal Editable As Boolean) As IValueEditor
    Dim Editor As EnumValueEditor
    Set Editor = New EnumValueEditor
    Call Editor.Initialize(Parent, FieldDefinition, Editable)
    Set Create = Editor
End Function

Friend Sub Initialize(ByVal Parent As Object, ByVal FieldDefinition As FieldDefinition, ByVal Editable As Boolean)
    Dim Item As Variant

    Set Definition    = FieldDefinition
    IsEditable        = Editable
    Set ControlValue  = Parent.Controls.Add("Forms.ComboBox.1")
    ControlValue.Name = "EnumEditor" & Definition.Name

    For Each Item In Definition.Validator.AllowedValues
        Call ControlValue.AddItem(Item)
    Next Item

    ControlValue.Locked = Not IsEditable
    If Not IsEditable Then ControlValue.BackColor = RGB(160,160,160)
End Sub

Private Property Get IValueEditor_Control() As MSForms.Control
    Set IValueEditor_Control = ControlValue
End Property

Private Property Get IValueEditor_Value() As Variant
    IValueEditor_Value = ControlValue.Value
End Property

Private Property Let IValueEditor_Value(ByVal Value As Variant)
    If IsNull(Value) Or IsEmpty(Value) Then
        ControlValue.ListIndex = -1
    Else
        ControlValue.Value = Value
    End If
End Property

Private Function IValueEditor_Validate() As Boolean
    IValueEditor_Validate = Definition.Validator.IsValid(DataEntry.Create(IValueEditor_Value))
End Function

Private Function IValueEditor_Convert() As Variant
    IValueEditor_Convert = Definition.Validator.Convert(DataEntry.Create(IValueEditor_Value))
End Function
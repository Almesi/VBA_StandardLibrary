VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ValueEditorFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Function Create(ByVal Parent As Object, ByVal Definition As FieldDefinition, ByVal Editable As Boolean) As IValueEditor
    Dim SearchFor As String
    SearchFor = TypeName(Definition.Validator)
    Select Case SearchFor
        Case "GenericValidator"
            Set Create = TextValueEditor.Create(Parent, Definition, Editable)
        Case "EnumValidator"
            Set Create = EnumValueEditor.Create(Parent, Definition, Editable)
        Case Else
            Err.Raise vbObjectError + 1000, "ValueEditorFactory", "Unsupported Validator: " & SearchFor
    End Select
End Function
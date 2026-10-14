VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FieldEntryForm
   Caption         =   "Update a single Value"
   ClientHeight    =   2000
   ClientWidth     =   4000
   OleObjectBlob   =   "FieldEntryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FieldEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IView

Private Type TView
    ReturnValue As FormReturn
    Model       As FieldEntryModel
End Type

Private This As TView

Private Editor    As IValueEditor
Private nEditable As Boolean

Public Function Create(ByVal Editable As Boolean) As FieldEntryForm
    Set Create = New FieldEntryForm
    Call Create.Initialize(Editable)
End Function

Friend Sub Initialize(ByVal Editable As Boolean)
    OKButton.TabIndex     = 2
    CancelButton.TabIndex = 3
    nEditable = Editable
End Sub

Private Sub UserForm_Initialize()
    This.ReturnValue = FormReturnCancel
End Sub

Public Sub LoadValue(ByVal Model As FieldEntryModel)
    Dim Definition As FieldDefinition

    Set This.Model = Model
    Set Definition = This.Model.Entry.Definition

    ClearDynamicControls

    Set Editor = ValueEditorFactory.Create(Me, Definition, nEditable)
    Call ConfigureEditor(Editor, 10, 100, 1)

    Editor.Value = This.Model.Entry.DataEntry.Value
End Sub

Private Sub ConfigureEditor(ByVal Editor As IValueEditor, ByVal LeftPosition As Single, ByVal TopPosition As Single, ByVal TabIndex As Long)
    With Editor.Control
        .Left     = LeftPosition
        .Top      = TopPosition
        .Width    = 100
        .Height   = 30
        .TabIndex = TabIndex
    End With
End Sub

Private Sub ClearDynamicControls()
    Dim Control As MSForms.Control
    Dim Index As Long

    For Index = Me.Controls.Count - 1 To 0 Step -1
        Set Control = Me.Controls.Item(Index)

        If Control.Name <> "ValueLabel" And Control.Name <> "OkButton" And Control.Name <> "CancelButton" Then
            Call Me.Controls.Remove(Control.Name)
        End If
    Next Index
End Sub

Private Sub OKButton_Click()
    If Not ValidateEditor Then Exit Sub

    SaveValue
    This.ReturnValue = FormReturnOK
    Me.Hide
End Sub

Private Sub CancelButton_Click()
    This.ReturnValue = FormReturnCancel
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.ReturnValue = FormReturnCancel
    Hide
End Sub

Private Function ValidateEditor() As Boolean
    If Not Editor.Validate() Then
        Editor.Control.SetFocus
        ValidateEditor = False
        Exit Function
    End If

    ValidateEditor = True
End Function

Private Sub SaveValue()
    This.Model.Entry.DataEntry.Value = Editor.Value
End Sub

Private Function IView_Show(ByVal Model As Object) As FormReturn
    Set This.Model = Model
    LoadValue This.Model
    Me.Show
    IView_Show = This.ReturnValue
End Function
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RowForm
   Caption         =   "Update Row"
   ClientHeight    =   5000
   ClientWidth     =   7000
   OleObjectBlob   =   "RowForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IView

Private Type TView
    ReturnValue As FormReturn
    Model       As RowModel
End Type

Private This As TView

Private Editors   As Collection
Private Labels    As Collection
Private nEditable As Boolean

Public Function Create(ByVal Editable As Boolean) As RowForm
    Set Create = New RowForm
    Call Create.Initialize(Editable)
End Function

Friend Sub Initialize(ByVal Editable As Boolean)
    Set Editors           = New Collection
    Set Labels            = New Collection
    OKButton.TabIndex     = 2
    CancelButton.TabIndex = 3
    nEditable             = Editable
End Sub

Private Sub UserForm_Initialize()
    This.ReturnValue = FormReturnCancel
    Set Editors      = New Collection
    Set Labels       = New Collection
End Sub

Public Sub LoadRow(ByVal Model As RowModel)
    Dim FieldEntry   As FieldEntry
    Dim Definition   As FieldDefinition
    Dim Editor       As IValueEditor
    Dim Label        As MSForms.Label
    Dim FieldIndex   As Long
    Dim LeftPosition As Single

    Set This.Model = Model
    Set Editors    = New Collection
    Set Labels     = New Collection

    ClearDynamicControls

    LeftPosition = 10

    For FieldIndex = 1 To This.Model.Row.Count
        Set FieldEntry = This.Model.Row.FieldEntries.Item(FieldIndex)
        Set Definition = FieldEntry.Definition

        Set Label = CreateLabel(Definition.Name, LeftPosition, 40)
        Call Labels.Add(Label)

        Set Editor = ValueEditorFactory.Create(Me, Definition, nEditable)
        Call ConfigureEditor(Editor, LeftPosition, 70, FieldIndex + 1)
        Editor.Value = FieldEntry.DataEntry.Value
        Call Editors.Add(Editor)

        LeftPosition = LeftPosition + 100
    Next FieldIndex

    OKButton.TabIndex = This.Model.Row.Count + 1
    CancelButton.TabIndex = This.Model.Row.Count + 2
    Me.Scrollwidth = LeftPosition + 200
End Sub

Private Function CreateLabel(ByVal Caption As String, ByVal LeftPosition As Single, ByVal TopPosition As Single) As MSForms.Label
    Dim Label As MSForms.Label

    Set Label     = Me.Controls.Add("Forms.Label.1")
    Label.Caption = Caption
    Label.Left    = LeftPosition
    Label.Top     = TopPosition
    Label.Width   = 100
    Label.Height  = 30

    Set CreateLabel = Label
End Function

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

        If Control.Name <> "OkButton" And Control.Name <> "CancelButton" Then
            Call Me.Controls.Remove(Control.Name)
        End If
    Next Index
End Sub

Private Sub OKButton_Click()
    If Not ValidateEditors Then Exit Sub

    SaveValues
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

Private Function ValidateEditors() As Boolean
    Dim Editor As IValueEditor

    For Each Editor In Editors
        If Not Editor.Validate() Then
            Editor.Control.SetFocus
            ValidateEditors = False
            Exit Function
        End If
    Next Editor

    ValidateEditors = True
End Function

Private Sub SaveValues()
    Dim FieldEntry As FieldEntry
    Dim Editor     As IValueEditor
    Dim FieldIndex As Long

    For FieldIndex = 1 To This.Model.Row.Count
        Set FieldEntry = This.Model.Row.FieldEntries.Item(FieldIndex)
        Set Editor = Editors.Item(FieldIndex)
        FieldEntry.DataEntry.Value = Editor.Convert()
    Next FieldIndex
End Sub

Private Function IView_Show(ByVal Model As Object) As FormReturn
    Set This.Model = Model
    LoadRow This.Model
    Me.Show
    IView_Show = This.ReturnValue
End Function
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListForm
   Caption         =   "Select Item"
   ClientHeight    =   4000
   ClientWidth     =   5000
   OleObjectBlob   =   "ListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IView

Private Type TView
    ReturnValue As FormReturn
    Model       As ListModel
End Type

Private This As TView

Private nMultiSelect As Boolean

Public Function Create(ByVal MultiSelect As Boolean) As ListForm
    Set Create = New ListForm
    Call Create.Initialize(MultiSelect)
End Function

Friend Sub Initialize(ByVal MultiSelect As Boolean)
    nMultiSelect = MultiSelect

    ListBox.MultiSelect = IIf(MultiSelect, fmMultiSelectMulti, fmMultiSelectSingle)
    ListLabel.Caption   = IIF(MultiSelect, Replace(ListLabel.Caption, "1 Item", "1 or more Items"), ListLabel.Caption)
End Sub

Private Sub UserForm_Initialize()
    This.ReturnValue = FormReturnCancel
End Sub

Private Function HasSelection() As Boolean
    Dim Index As Long

    If nMultiSelect Then
        For Index = 0 To ListBox.ListCount - 1
            If ListBox.Selected(Index) Then
                HasSelection = True
                Exit Function
            End If
        Next Index
    Else
        HasSelection = ListBox.ListIndex >= 0
    End If
End Function

Private Function InitializeList()
    Dim Item As Variant
    For Each Item In This.Model.Items
        Call ListBox.AddItem(CStr(Item))
    Next Item

    ListBox.Height = This.Model.Count * 10
    Me.ScrollHeight = ListBox.Height
End Function

Private Sub SaveSelection()
    Dim Index As Long
    Dim Selected As Collection
    Set Selected = This.Model.Selected

    If nMultiSelect Then
        For Index = 0 To ListBox.ListCount - 1
            If ListBox.Selected(Index) Then
                Call Selected.Add(Index + 1) ' To fit Collection 1 Base
            End If
        Next Index
    ElseIf ListBox.ListIndex >= 0 Then
        Call Selected.Add(ListBox.ListIndex)
    End If
End Sub

Private Sub OKButton_Click()
    If Not HasSelection Then Exit Sub

    Call SaveSelection

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

Private Function IView_Show(ByVal Model As Object) As FormReturn
    Set This.Model = Model
    Call InitializeList()
    Me.Show
    IView_Show = This.ReturnValue
End Function
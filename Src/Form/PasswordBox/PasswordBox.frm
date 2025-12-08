VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PasswordBox 
   Caption         =   "PasswordBox"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "PasswordBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PasswordBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


Implements IView

Private Type TView
    IsCancelled As Boolean
    Model As SingleInputModel
End Type

Private This As TView

Private Function IView_Run(ByVal viewModel As Object, Optional LabelText As String = Empty) As Boolean
    Set This.Model = viewModel
    Label.Caption = LabelText
    Show
    IView_Run = Not This.IsCancelled ' cancelled
End Function

Private Sub Password_Change()
    This.Model.Value = Password.Text
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OkButton_Click()
    Hide
End Sub
 
Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Hide
End Sub
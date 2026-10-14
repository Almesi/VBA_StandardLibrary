VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YesNoForm 
   Caption         =   "Message"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9450.001
   OleObjectBlob   =   "YesNoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "YesNoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IView

Private Type TView
    ReturnValue As FormReturn
End Type

Private This As TView

Public Function Create(ByVal Question As String) As YesNoForm
    Set Create = New YesNoForm
    Call Create.Initialize(nCallable)
End Function

Friend Sub Initialize(ByVal Question As String)
    Me.Caption = Question
    MessageLabel.Caption = Question
End Sub

Private Sub YesButton_Click()
    This.ReturnValue = FormReturnYes
    Me.Hide
End Sub

Private Sub CancelButton_Click()
    This.ReturnValue = FormReturnNo
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
    Show
    IView_Show = This.ReturnValue
End Function

Private Property Let IView_Result(ByVal Value As FormReturn)
    This.ReturnValue = Value
End Property

Private Property Get IView_Form() As Object
    Set IView_Form = Me
End Property
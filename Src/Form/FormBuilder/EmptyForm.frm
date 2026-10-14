VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmptyForm 
   Caption         =   "EmptyForm"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "EmptyForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EmptyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IView


Private Type TView
    ReturnValue As FormReturn
    Model       As Object
End Type

Private This As TView

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = FormReturnCancel
        OnCancel
    End If
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.ReturnValue = FormReturnCancel
    Hide
End Sub



Private Function IView_Show(ByVal Model As Object) As FormReturn
    Set This.Model = Model
    Show
    IView_Show = This.ReturnValue
End Function

Private Property Let IView_Result(ByVal Value As FormReturn)
    This.ReturnValue = Value
End Property

Private Property Get IView_Form() As Object
    Set IView_Form = Me
End Property
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FreeValueForm 
   Caption         =   "Change a single Value"
   ClientHeight    =   2505
   ClientWidth     =   2370
   OleObjectBlob   =   "FreeValueForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FreeValueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IView

Private Type TView
    ReturnValue As FormReturn
    Model       As ValueModel
End Type

Private This As TView

Public Function Create() As FreeValueForm
    Set Create = New FreeValueForm
End Function

Private Sub OKButton_Click()
    This.Model.Value = ValueBox.Value
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
    Show
    IView_Show = This.ReturnValue
End Function

Private Property Let IView_Result(ByVal Value As FormReturn)
    This.ReturnValue = Value
End Property

Private Property Get IView_Form() As Object
    Set IView_Form = Me
End Property
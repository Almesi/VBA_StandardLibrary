VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MessageForm 
   Caption         =   "Message"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9450.001
   OleObjectBlob   =   "MessageForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IView

Public Enum MessageType
    MessageInformation = 0
    MessageWarning     = 1
    MessageError       = 2
End Enum


Private Type TView
    ReturnValue As FormReturn
End Type

Private This As TView

Public Function Create(ByVal Title As String, ByVal Message As String, ByVal MessType As MessageType) As MessageForm
    Set Create = New MessageForm
    Call Create.Initialize(nCallable)
End Function

Friend Sub Initialize(ByVal Title As String, ByVal Message As String, ByVal MessType As MessageType)
    Me.Caption = Title
    MessageLabel.Caption = Message

    InfoImage.Visible    = (MessType = MessageType.MessageInformation)
    WarningImage.Visible = (MessType = MessageType.MessageWarning)
    ErrorImage.Visible   = (MessType = MessageType.MessageError)
End Sub

Private Sub OKButton_Click()
    This.ReturnValue = FormReturnOK
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.ReturnValue = FormReturnOK
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
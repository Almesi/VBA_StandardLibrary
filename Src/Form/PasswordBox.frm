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
Attribute VB_Exposed = False

Option Explicit

Private CloseWindow As Boolean
Private SavedPassword As String
Private p_SavePassword As Boolean


Public Function Enter(LabelText As String) As String
    If SavedPassword <> Empty Then
        Enter = SavedPassword
    Else
    Me.Label.Caption = LabelText
    Me.Show
    If CloseWindow Then Password.Text = Empty
    If p_SavePassword Then SavedPassword = Password.Text
    Enter = Password.Text

    Password.Text = Empty
    CloseWindow = False
    Me.Label.Caption = Empty
    End If
End Function

Public Function Check(LabelText As String, Text As String) As Boolean
    Check = Enter(LabelText) Like Text
End Function

Public Property Let SavePassword(Value As Boolean)
    p_SavePassword = Value
End Property


Private Sub Password_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
        KeyCode = 0
        If KeyCode = vbKeyEscape Then CloseWindow = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    CloseWindow = True
End Sub

Private Sub CancelButton_Click()
    CloseWindow = True
    Me.Hide
End Sub

Private Sub OKButton_Click()
    Me.Hide
End Sub

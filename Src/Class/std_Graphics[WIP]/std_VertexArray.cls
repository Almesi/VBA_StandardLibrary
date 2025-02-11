VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_VertexArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private p_ID As Long

Private Sub Class_Initialize()
    GLCall(glGenVertexArrays(1, VarPtr(p_ID)))
End Sub

Private Sub Class_Terminate()
    GLCall(glDeleteVertexArrays(1, VarPtr(p_ID)))
End Sub

Public Sub AddBuffer(VB As VertexBuffer, Layout As VertexBufferLayout)
    Dim Elements As VertexBufferLayout.VertexBufferElement
    Dim Element As VertexBufferLayout.VertexBufferElement
    Dim Offset As Long
    Dim i As Long

    Call Bind()
    Call VB.Bind()
    Elements = Layout.Elements
    For i = 0 To Elements.Size
        Element = Elements(i)
        With Elements(i)
            Call GLCall(glEnableVertexAttribArray(i))
            Call GLCall(glVertexAttribPointer(i, .Count, .Typee, .Normalized, Layout.GetStride(), INT2VOIDP(Offset)))
            Offset += .Count * .GetSizeOfType(.Typee);
        End With
    Next i

End Sub

Public Sub Bind()
    Call GLCall(glBindVertexArray(p_ID))
End Sub

Public Sub Bind()
    Call GLCall(glBindVertexArray(0))
End Sub
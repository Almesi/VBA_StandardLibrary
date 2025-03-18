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
    Call glGenVertexArrays(1, p_ID)
End Sub

Private Sub Class_Terminate()
    Call glDeleteVertexArrays(1, p_ID)
End Sub

Public Sub AddBuffer(VB As std_Buffer, Layout As std_BufferLayout)
    Dim Offset As Long
    Dim i As Long

    Call Bind()
    Call VB.Bind()
    For i = 0 To Ubound(Layout.Elements)
        With Layout.Element(i)
            Call glEnableVertexAttribArray(i)
            Call glVertexAttribPointer(i, .Count, .Typee, .Normalized, Layout.Stride(), Offset)
            Offset = Offset + (.Count * Layout.GetSizeOfType(.Typee))
        End With
    Next i
End Sub

Public Sub Bind()
    Call glBindVertexArray(p_ID)
End Sub

Public Sub UnBind()
    Call glBindVertexArray(0)
End Sub
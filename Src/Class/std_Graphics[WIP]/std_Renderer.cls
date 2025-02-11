VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Renderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Public Sub Clear()
    Call GLCall(glClear( GL_COLOR_BUFFER_BIT))
End Sub

Public Sub Draw(VA As std_VertexArray, IB As std_IndexBuffer, Shader As std_Shader) const
    Call Shader.Bind()
    Call VA.Bind()
    Call IB.Bind()
    Call GLCall(glDrawElements(GL_TRIANGLES, IB.Count, GL_UNSIGNED_INT, nullptr))
End Sub
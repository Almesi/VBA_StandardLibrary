VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_IndexBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private p_ID As Long
Private p_Count As Long

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Call GLCall(glDeleteBuffers(1, VarPtr(p_ID)) );
End Sub

Public Property Get ID() As Long
    ID = p_ID
End Property

Public Property Get IDAdress() As LongPtr
    IDAdress = VarPtr(p_ID)
End Property

Public Property Get Count As Long
    Count = p_Count
End Property

Private Function Create(Indices As Long, Count As Long) As std_IndexBuffer
    Set Create = New std_IndexBuffer
    With Create
        ASSERT(LenB(p_ID) == sizeof(GLuint))
        Call GLCall(glGenBuffers(1, VarPtr(.IDAdress)))
        Call GLCall(glBindBuffer(GL_ELEMENT_ARRAY_BUFFER, .ID))
        Call GLCall(glBufferData(GL_ELEMENT_ARRAY_BUFFER, Count * LenB(p_ID), Indices, GL_STATIC_DRAW))
    End With
End Function

Public Sub Bind()
    Call GLCall(glBindBuffer(GL_ELEMENT_ARRAY_BUFFER, p_ID))
End Sub

Public Sub Unbind()
    Call GLCall(glBindBuffer(GL_ELEMENT_ARRAY_BUFFER, 0))
End Sub
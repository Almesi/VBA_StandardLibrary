VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_VertexBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit



Private p_ID As Long

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Call glDeleteBuffers(1, CLng(VarPtr(p_ID)))
End Sub

Public Property Get IDAdress() As LongPtr
    IDAdress = VarPtr(p_ID)
End Property

Public Property Get ID() As Long
    ID = p_ID
End Property

Public Function Create(Vertices() As Double) As std_VertexBuffer
    Set Create = New std_VertexBuffer
    With Create
        Call glGenBuffers(1, .IDAdress)
        Call glBindBuffer(GL_ARRAY_BUFFER, .ID)
        Call glBufferData(GL_ARRAY_BUFFER, Ubound(Vertices) * LenB(Vertices(0)), Vertices, GL_STATIC_DRAW)
    End With
End Function

Public Sub Bind()
    Call glBindBuffer(GL_ARRAY_BUFFER, p_ID)
End Sub

Public Sub UnBind()
    Call glBindBuffer(GL_ARRAY_BUFFER, 0)
End Sub
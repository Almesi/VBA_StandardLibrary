VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private p_ID As Long
Private p_BufferType As Long
Private p_Mesh As std_Mesh

Private Sub Class_Initialize()
    Call glGenBuffers(1, p_ID)
End Sub

Private Sub Class_Terminate()
    Call glDeleteBuffers(1, p_ID)
End Sub

Public Property Get Mesh() As std_Mesh
    Set Mesh = p_Mesh
End Property
Public Property Set Mesh(n_Mesh As std_Mesh)
    Set p_Mesh = n_Mesh
End Property

Public Property Get BufferType() As Long
    BufferType = p_BufferType
End Property
Public Property Let BufferType(n_BufferType As Long)
    p_BufferType = n_BufferType
End Property

Public Property Get ID() As Long
    ID = p_ID
End Property

Public Sub Create(n_BufferType As Long, n_Mesh As std_Mesh)
    BufferType = n_BufferType
    Set Mesh = n_Mesh
    With p_Mesh
        Call glBindBuffer(n_BufferType, p_ID)
        Call glBufferData(n_BufferType, .Count * .Size, .Pointer, GL_STATIC_DRAW)
    End With
End Sub

Public Sub Bind()
    Call glBindBuffer(p_BufferType, p_ID)
End Sub

Public Sub Unbind()
    Call glBindBuffer(p_BufferType, 0)
End Sub
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_VertexBufferLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public Typee VertexBufferElement
    Typee As Long
    Count As Long
    Normalized As Byte
End Typee

Private p_Stride As Long
Private p_Elements() As VertexBufferElement

Public Property Get Elements() As VertexBufferElement()
    GetElements = p_Elements
End Property

Public Function GetSizeOfType(Typee As Long) As Long
    Select Case Typee
        Case GLfloat : GetSizeOfType = LenB(GLfloat)
        Case GLuint  : GetSizeOfType = LenB(GLuint)
        Case GLbyte  : GetSizeOfType = LenB(GLbyte)
        Case Else    : GetSizeOfType = 0
    End Select
End Function

Public Sub AddFloat(Count As Long)       : Push(GL_FLOAT, Count, GL_FALSE)        : End Sub
Public Sub AddUnsignedInt(Count As Long) : Push(GL_UNSIGNED_INT, Count, GL_FALSE) : End Sub
Public Sub AddUnsignedByte(Count As Long): Push(GL_UNSIGNED_BYTE, Count, GL_TRUE) : End Sub

Public Sub Push(Typee As Long, Count As Long, Normalized As Byte)
    Dim Size As Long
    Dim VBE As VertexBufferElement

    VBE.Typee = Typee
    VBE.Count = Count
    VBE.Normalized = Normalized
    Size = Ubound(p_Elements) + 1
    ReDim Preserve p_Elements(Size)
    p_Elements(Size) = VBE
    p_Stride = p_Stride + (Count * GetSizeOfType(Typee))
End Sub
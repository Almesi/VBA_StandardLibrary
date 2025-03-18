VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_BufferLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private p_Stride As Long
Private p_Elements() As New std_BufferElement

Public Property Get Stride() As Long
    Stride = p_Stride
End Property

Public Property Get Elements() As std_BufferElement()
    Elements = p_Elements
End Property

Public Property Get Element(Index As Long) As std_BufferElement
    Set Element = p_Elements(Index)
End Property

Public Function GetSizeOfType(Typee As Long) As Long
    Select Case Typee
        Case GL_DOUBLE        : GetSizeOfType = 8
        Case GL_FLOAT         : GetSizeOfType = 4
        Case GL_UNSIGNED_INT  : GetSizeOfType = 4
        Case GL_UNSIGNED_BYTE : GetSizeOfType = 1
        Case Else             : GetSizeOfType = 0
    End Select
End Function

Public Sub AddDouble(Count As Long)      : Call Push(GL_DOUBLE, Count, GL_FALSE)       : End Sub
Public Sub AddFloat(Count As Long)       : Call Push(GL_FLOAT, Count, GL_FALSE)        : End Sub
Public Sub AddUnsignedInt(Count As Long) : Call Push(GL_UNSIGNED_INT, Count, GL_FALSE) : End Sub
Public Sub AddUnsignedByte(Count As Long): Call Push(GL_UNSIGNED_BYTE, Count, GL_TRUE) : End Sub

Public Sub Push(Typee As Long, Count As Long, Normalized As Byte)
    Dim Size As Long
    Dim VBE As New std_BufferElement
    On Error Resume Next

    VBE.Typee = Typee
    VBE.Count = Count
    VBE.Normalized = Normalized
    Size = Ubound(p_Elements) + 1
    ReDim Preserve p_Elements(Size)
    Set p_Elements(Size) = VBE
    p_Stride = p_Stride + (Count * GetSizeOfType(Typee))
End Sub
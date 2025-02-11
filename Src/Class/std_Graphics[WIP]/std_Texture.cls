VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Texture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private p_ID As Long
Private p_FilePath As String
Private p_LocalBuffer() As Byte
Private p_Width As Long
Private p_Height As Long
Private p_BPP As Long

Public Function Create(Path As String) As std_Texture
    p_FilePath = Path
    Call stbi_set_flip_vertically_on_load(1)

    p_LocalBuffer = stbi_load(VarPtr(Path), VarPtr(p_Width), VarPtr(p_Height), VarPtr(p_BPP), 4)
    Call GLCall(glGenTextures(1, VarPtr(p_ID)))
    Call GLCall(glBindTexture(GL_TEXTURE_2D, p_ID))

    Call GLCall(glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR))
    Call GLCall(glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR))
    Call GLCall(glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE))
    Call GLCall(glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE))

    Call GLCall(glTexImage2D(GL_TEXTURE_2D, 0, GL_RGBA8, p_Width, p_Height, 0, GL_RGBA, GL_UNSIGNED_BYTE, p_LocalBuffer))
    Unbind()

    If Ubound(p_LocalBuffer) >= 0 Then stbi_image_free(p_LocalBuffer);
End Function

Public Sub Bind(Slot As Long)
    Call GLCall(glActiveTexture(GL_TEXTURE0 + Slot))
    Call GLCall(glBindTexture(GL_TEXTURE_2D, p_ID))
End Sub

Public Sub Unbind()
    Call GLCall(glBindTexture(GL_TEXTURE_2D, 0))
End Sub

Private Sub Class_Terminate()
    Call GLCall(glDeleteTextures(1, VarPtr(p_ID)))
End Sub
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

Private Sub Class_Initialize()
    Call glGenTextures(1, p_ID)
End Sub

Private Sub Class_Terminate()
    Call glDeleteTextures(1, p_ID)
End Sub

Public Function Create(Path As String) As std_Texture
    Dim Image As New WIA.ImageFile

    Call Image.LoadFile(Path)
    p_LocalBuffer = Image.FileData.BinaryData
    p_Width = Image.Width
    p_Height = Image.Height
    p_BPP = 4
    p_FilePath = Path
    
    Call Bind(0)
    Call glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR)
    Call glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)
    Call glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE)
    Call glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE)
    Call glTexImage2D(GL_TEXTURE_2D, 0, GL_RGBA8, p_Width, p_Height, 0, GL_RGBA, GL_UNSIGNED_BYTE, VarPtr(p_LocalBuffer(0)))
End Function

Public Sub Bind(Slot As Long)
    Call glActiveTexture(GL_TEXTURE0 + Slot)
    Call glBindTexture(GL_TEXTURE_2D, p_ID)
End Sub

Public Sub Unbind()
    Call glBindTexture(GL_TEXTURE_2D, 0)
End Sub

Sub GetImageInfoWIA()
    Dim imgPath As String
    Dim img As WIA.ImageFile
    Dim width As Integer, height As Integer
    Dim i As Long
    
    ' Create WIA ImageFile object
    Set img = New WIA.ImageFile
    
    
    imgPath = "C:\Users\deallulic\Documents\TestFolder\TestTexture.png"
    img.LoadFile imgPath
    
    ' Get dimensions
    width = img.width
    height = img.height
    
    Dim one As LongPtr, two As LongPtr
    Dim bytes() As Byte
    bytes() = img.FileData.BinaryData
    
    For i = 0 To 40000
        one = VarPtr(bytes(i))
        two = VarPtr(bytes(0))
        Debug.Print one - two
    Next i
End Sub

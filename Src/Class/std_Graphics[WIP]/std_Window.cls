VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private p_ID As Long

Private Sub Class_Initialize()
End Sub

Public Property Get ID() As Long
    ID = p_ID
End Property

Private Sub Class_Terminate()
    Call glutDestroyWindow(p_ID)
End Sub

Public Function Create(Width As Long, _
                       Height As Long, _
                       DisplayMode As Long, _
                       Optional Title As String = Empty, _
                       Optional Version As String = "4_6", _
                       Optional ContextProfile As Long = 0, _
                       Optional ContextFlag As Long = 0) _
                       As Long
    Dim VersionIndex() As String
    Dim Major As Long
    Dim Minor As Long

    VersionIndex = Split(Version, "_")
    If Ubound(VersionIndex) < 0 Then
        Exit Function
    ElseIf Ubound(VersionIndex) = 0 Then
        Major = 4
        Minor = 6
    Else
        Major = CLng(VersionIndex(0))
        Minor = CLng(VersionIndex(1))
    End If

    Call glutInitContextVersion(Major, Minor)
    Call glutInitContextProfile(ContextProfile)

    Call glutInitWindowSize(Width, Height)
    Call glutInitDisplayMode(DisplayMode)
    Call glutSetOption(GLUT_ACTION_ON_WINDOW_CLOSE, GLUT_ACTION_GLUTMAINLOOP_RETURNS)
    p_ID = glutCreateWindow("OpenGL Test")


    Call RemapVBToGL(Version)
    Create = p_ID
End Function

Public Function GetWindowHwnd(Window As Long)
    GetWindowHwnd = glutGetWindowHwnd(Window)
End Function

Public Function RemapVBToGL(Optional ToInclude As String = "ALL") As Boolean
    Repeat:
    Select Case ToInclude
    Case "ALL"  :                                         : ToInclude = "OTHER": GoTo Repeat
    Case "OTHER": Call Application.Run("RemapVBToGLOTHER"): ToInclude = "WGL"  : GoTo Repeat
    Case "WGL"  : Call Application.Run("RemapVBToGLWGL")  : ToInclude = "EXT"  : GoTo Repeat
    Case "EXT"  : Call Application.Run("RemapVBToGLEXT")  : ToInclude = "ARB"  : GoTo Repeat
    Case "ARB"  : Call Application.Run("RemapVBToGLARB")  : ToInclude = "4_6"  : GoTo Repeat
    Case "4_6"  : Call Application.Run("RemapVBToGL4_6")  : ToInclude = "4_5"  : GoTo Repeat
    Case "4_5"  : Call Application.Run("RemapVBToGL4_5")  : ToInclude = "4_4"  : GoTo Repeat
    Case "4_4"  : Call Application.Run("RemapVBToGL4_4")  : ToInclude = "4_3"  : GoTo Repeat
    Case "4_3"  : Call Application.Run("RemapVBToGL4_3")  : ToInclude = "4_2"  : GoTo Repeat
    Case "4_2"  : Call Application.Run("RemapVBToGL4_2")  : ToInclude = "4_1"  : GoTo Repeat
    Case "4_1"  : Call Application.Run("RemapVBToGL4_1")  : ToInclude = "4_0"  : GoTo Repeat
    Case "4_0"  : Call Application.Run("RemapVBToGL4_0")  : ToInclude = "3_3"  : GoTo Repeat
    Case "3_3"  : Call Application.Run("RemapVBToGL3_3")  : ToInclude = "3_2"  : GoTo Repeat
    Case "3_2"  : Call Application.Run("RemapVBToGL3_2")  : ToInclude = "3_1"  : GoTo Repeat
    Case "3_1"  : Call Application.Run("RemapVBToGL3_1")  : ToInclude = "3_0"  : GoTo Repeat
    Case "3_0"  : Call Application.Run("RemapVBToGL3_0")  : ToInclude = "2_1"  : GoTo Repeat
    Case "2_1"  : Call Application.Run("RemapVBToGL2_1")  : ToInclude = "2_0"  : GoTo Repeat
    Case "2_0"  : Call Application.Run("RemapVBToGL2_0")  : ToInclude = "1_5"  : GoTo Repeat
    Case "1_5"  : Call Application.Run("RemapVBToGL1_5")  : ToInclude = "1_4"  : GoTo Repeat
    Case "1_4"  : Call Application.Run("RemapVBToGL1_4")  : ToInclude = "1_3"  : GoTo Repeat
    Case "1_3"  : Call Application.Run("RemapVBToGL1_3")  : ToInclude = "1_2"  : GoTo Repeat
    Case "1_2"  : Call Application.Run("RemapVBToGL1_2")  : ToInclude = "1_1"  : GoTo Repeat
    Case "1_1"  : Call Application.Run("RemapVBToGL1_1")  : RemapVBToGL = True
    Case Else
    End Select
End Function
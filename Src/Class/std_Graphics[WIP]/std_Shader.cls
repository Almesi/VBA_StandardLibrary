VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Shader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Const GL_VERTEX_SHADER   As Long = &H8B31
Private Const GL_FRAGMENT_SHADER As Long = &H8B30
Private Const GL_COMPILE_STATUS  As Long = &H8B81
Private Const GL_INFO_LOG_LENGTH As Long = &H8B84

Private p_VertexShader As Long
Private p_FragmentShader As Long
Private p_LinkedShader As Long



Public Property Get VertexShader()          : VertexShader     = p_VertexShader   : End Property
Public Property Get FragmentShader()        : FragmentShader   = p_FragmentShader : End Property
Public Property Get LinkedShader()          : LinkedShader     = p_LinkedShader   : End Property

Public Property Let VertexShader(n_Value)   : p_VertexShader   = n_Value          : End Property
Public Property Let FragmentShader(n_Value) : p_FragmentShader = n_Value          : End Property
Public Property Let LinkedShader(n_Value)   : p_LinkedShader   = n_Value          : End Property


Public Function CreateFromFile(VertexShaderFile As String, FragmentShaderFile As String) As std_Shader
    Dim FileNumber As Long
    Dim VertexShaderSource As String
    Dim FragmentShaderSource As String
    Dim Line As String

    FileNumber = FreeFile
    Open VertexShaderFile For Input As #FileNumber
    Do Until EOF(FileNumber)
        Line Input #FileNumber, Line
        VertexShaderSource = VertexShaderSource & Line
    Loop
    Close #FileNumber
    
    FileNumber = FreeFile
    Open FragmentShaderFile For Input As #FileNumber
    Do Until EOF(FileNumber)
        Line Input #FileNumber, Line
        FragmentShaderSource = FragmentShaderSource & Line
    Loop
    Close #FileNumber
    Set CreateFromFile = CreateFromString(VertexShaderSource, FragmentShaderSource)
End Function

Public Function CreateFromString(VertexShaderSource As String, FragmentShaderSource As String) As std_Shader
    Dim SourcePtr As LongPtr

    Set CreateFromString = New std_Shader

    With CreateFromString 

        .VertexShader   = glCreateShader(GL_VERTEX_SHADER)
        .FragmentShader = glCreateShader(GL_FRAGMENT_SHADER)

        SourcePtr = VarPtr(VertexShaderSource)
        Call glShaderSource(.VertexShader, 1, SourcePtr, Len(VertexShaderSource))
        SourcePtr = VarPtr(VertexShaderSource)
        Call glShaderSource(.FragmentShader, 1, SourcePtr, Len(FragmentShaderSource))

        Call glCompileShader(.VertexShader)
        Call glCompileShader(.FragmentShader)

        If .CompileStatus(.VertexShader)   = False Then Exit Function
        If .CompileStatus(.FragmentShader) = False Then Exit Function
        .LinkedShader = glCreateProgram()
        Call glAttachShader(.LinkedShader, .VertexShader)
        Call glAttachShader(.LinkedShader, .FragmentShader)
        Call glLinkProgram(.LinkedShader)
        If .LinkStatus(.LinkedShader) = False Then Exit Function  
    End With
End Function

Public Function CompileStatus(Shader As Long) As Boolean
    Dim Compiled As Long
    Call glGetShaderiv(Shader, GL_COMPILE_STATUS, Compiled)
    If Compiled = 0 Then
        GetError(Shader)
    Else
        CompileStatus = True
    End If
End Function

Public Function LinkStatus(Shader As Long) As Boolean
    Dim Linked As Long
    Call glGetProgramiv(Shader, GL_LINK_STATUS, Linked)
    If Linked = 0 Then
        Debug.Print GetError(Shader)
    Else
        CompileStatus = True
    End If
End Function

Private Function GetError(Shader As Long) As String
    Dim Log() As Byte
    Dim InfoLogLength As Long

    Call glGetProgramiv(Shader, GL_INFO_LOG_LENGTH, InfoLogLength)
    ReDim Log(InfoLogLength - 1)
    Call glGetProgramInfoLog(Shader, InfoLogLength, 0, Log(0))
    GetError = StrConv(Log, vbUnicode)
End Function

Private Sub DeleteShader(Shader As Long)
    Shader = 0
End Sub

Private Sub DeleteShaders()
    Call DeleteShader(p_VertexShader)
    Call DeleteShader(p_FragmentShader)
    Call DeleteShader(p_LinkedShader)
End Sub

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Call DeleteShaders
End Sub
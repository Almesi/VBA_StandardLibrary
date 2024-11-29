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

Private Initialized As Boolean
Private ErrorCataLog(1, 99) As Variant
Private Handler As New std_Error

Private Const GL_VERTEX_SHADER As Long = &H8B31
Private Const GL_FRAGMENT_SHADER As Long = &H8B30
Private Const GL_COMPILE_STATUS As Long = &H8B81
Private Const GL_INFO_LOG_LENGTH As Long = &H8B84
Private Shader(2) As Long ' 0 = VertexShader, 1 = FragmentShader, 2 = Linked Shader

' Public Errormethods

    Public Property Let SetHandler(n_Handler As std_Error)
        Set Handler = n_Handler
    End Function

    Public Property Get IS_ERROR() As Boolean
        IS_ERROR = Handler.IS_ERROR
    End Property

    Public Function CreateFromString(VertexShaderSource As String, FragmentShaderSource As String, Optional ThrowError As Boolean = True) As Boolean

        Dim SourcePtr As LongPtr

        Shader(0) = glCreateShader(GL_VERTEX_SHADER)
        Shader(1) = glCreateShader(GL_FRAGMENT_SHADER)

        SourcePtr = VarPtr(VertexShaderSource)
        Call glVertexShaderSource(Shader(0), 1, SourcePtr, Len(VertexShaderSource))
        SourcePtr = VarPtr(VertexShaderSource)
        Call glVertexShaderSource(Shader(1), 1, SourcePtr, Len(FragmentShaderSource))

        Call glCompileShader(Shader(0))
        Call glCompileShader(Shader(1))

        If CompileStatus(0, ThrowError) = Handler.IS_ERROR Then CreateFromString = Handler.IS_ERROR: Exit Function
        If CompileStatus(1, ThrowError) = Handler.IS_ERROR Then CreateFromString = Handler.IS_ERROR: Exit Function
        Shader(2) = glCreateProgram()
        Call glAttachShader(LinkedShader, Shader(0))
        Call glAttachShader(LinkedShader, Shader(1))
        Call glLinkProgram(LinkedShader)
        CreateFromString = LinkStatus(2, ThrowError)
        
    End Function

    Public Function CreateFromFile(VertexShaderFile As String, FragmentShaderFile As String, Optional ThrowError As Boolean = True) As Boolean
        
        Dim FileNumber As Long
        Dim VertexShaderSource As String
        Dim FragmentShaderSource As String

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
        CreateFromFile = CreateFromString(VertexShaderSource, FragmentShaderSource, ThrowError)

    End Function

    Public Function GetShader(ShaderType As Long, Optional ThrowError As Boolean = True) As Long
        If CheckShaderType(ShaderType, ThrowError) <> Handler.IS_ERROR Then
            GetShader = Shader(ShaderType)
        Else
            GetShader = -1
        End If
    End Function

    Public Function DeleteShader(ShaderType As Long, Optional ThrowError As Boolean = True) As Boolean
        DeleteShader = CheckShaderType(ShaderType, ThrowError)
        If DeleteShader <> Handler.IS_ERROR Then
            If Shader(ShaderType) <> 0 Then
                Call glDeleteShader(Shader(ShaderType))
                Shader(ShaderType) = 0
            End If
        End If
    End Function

    Public Function DeleteShaders(Optional ThrowError As Boolean = True) As Boolean
        Dim i As Long
        For i = 0 To Ubound(Shader)
            DeleteShaders = DeleteShader(i, ThrowError)
            If DeleteShaders = Handler.IS_ERROR Then Exit Function
        Next i
    End Function

    Private Function CompileStatus(Index As Long) As Boolean
        Dim Compiled As Long
        Dim InfoLog As String
        Dim Log() As Byte
        Dim InfoLogLength As Long
        
        Call glGetShaderiv(Shader(Index), GL_COMPILE_STATUS, Compiled)
        If Compiled = 0 Then
            Call glGetShaderiv(Shader(Index), GL_INFO_LOG_LENGTH, InfoLogLength)
            ReDim Log(InfoLogLength - 1)
            Call glGetShaderInfoLog(Shader(Index), InfoLogLength, ByVal 0&, Log(0))
            InfoLog = StrConv(Log, vbUnicode)
            If ThrowError = False Then Exit Function
            CheckShaderType = Handler.Handle(ErrorCataLog, 2, ThrowError, Index, InfoLog)
        End If
    End Function

    Private Function LinkStatus(Index As Long) As Boolean
        Dim Linked As Long
        Dim InfoLog As String
        Dim Log() As Byte
        Dim InfoLogLength As Long

        Call glGetProgramiv(Shader(Index), GL_LINK_STATUS, Linked)
        If Linked = 0 Then
            Call glGetProgramiv(Shader(Index), GL_INFO_LOG_LENGTH, InfoLogLength)
            ReDim Log(InfoLogLength - 1)
            Call glGetProgramInfoLog(Shader(Index), InfoLogLength, ByVal 0&, Log(0))
            InfoLog = StrConv(Log, vbUnicode)
            If ThrowError = False Then Exit Function
            CheckShaderType = Handler.Handle(ErrorCataLog, 3, ThrowError, InfoLog)
        End If
    End Function

    Private Function CheckShaderType(ShaderType As Long, Optional ThrowError As Boolean = True) As Boolean
        If ShaderType > Ubound(Shader) Or ShaderType < Ubound(Shader) Then
            CheckShaderType = Handler.Handle(ErrorCataLog, 1, ThrowError, ShaderType)
        End If
    End Function


    Private Function Class_Initialize()
        ProtInit
    End Function

    Private Function Class_Terminate()
        Call DeleteShaders
    End Function

    
    Private Function ProtInit()
        If Initialized = False Then
            ErrorCataLog(0, 0000) = 0002: ErrorCataLog(1, 0000) = "std_Shader"
            ErrorCataLog(0, 0001) = 1001: ErrorCataLog(1, 0001) = "ShaderType doesnt exist"
            ErrorCataLog(0, 0002) = 1001: ErrorCataLog(1, 0002) = "Shader compilation failed"
            ErrorCataLog(0, 0003) = 1001: ErrorCataLog(1, 0003) = "Shader Linking failed"
            ErrorCataLog(0, 0004) = 1000: ErrorCataLog(1, 0004) = "PLACEHOLDER"
            ErrorCataLog(0, 0005) = 1000: ErrorCataLog(1, 0005) = "PLACEHOLDER"
            ErrorCataLog(0, 0006) = 1000: ErrorCataLog(1, 0006) = "PLACEHOLDER"
            ErrorCataLog(0, 0007) = 1000: ErrorCataLog(1, 0007) = "PLACEHOLDER"
            ErrorCataLog(0, 0008) = 1000: ErrorCataLog(1, 0008) = "PLACEHOLDER"
            ErrorCataLog(0, 0009) = 1000: ErrorCataLog(1, 0009) = "PLACEHOLDER"
            ErrorCataLog(0, 0010) = 1000: ErrorCataLog(1, 0010) = "PLACEHOLDER"
            ErrorCataLog(0, 0011) = 1000: ErrorCataLog(1, 0011) = "PLACEHOLDER"
            ErrorCataLog(0, 0012) = 1000: ErrorCataLog(1, 0012) = "PLACEHOLDER"
            ErrorCataLog(0, 0013) = 1000: ErrorCataLog(1, 0013) = "PLACEHOLDER"
            ErrorCataLog(0, 0014) = 1000: ErrorCataLog(1, 0014) = "PLACEHOLDER"
            ErrorCataLog(0, 0015) = 1000: ErrorCataLog(1, 0015) = "PLACEHOLDER"
            ErrorCataLog(0, 0016) = 1000: ErrorCataLog(1, 0016) = "PLACEHOLDER"
            ErrorCataLog(0, 0017) = 1000: ErrorCataLog(1, 0017) = "PLACEHOLDER"
            ErrorCataLog(0, 0018) = 1000: ErrorCataLog(1, 0018) = "PLACEHOLDER"
            ErrorCataLog(0, 0019) = 1000: ErrorCataLog(1, 0019) = "PLACEHOLDER"
            ErrorCataLog(0, 0020) = 1000: ErrorCataLog(1, 0020) = "PLACEHOLDER"
            ErrorCataLog(0, 0021) = 1000: ErrorCataLog(1, 0021) = "PLACEHOLDER"
            ErrorCataLog(0, 0022) = 1000: ErrorCataLog(1, 0022) = "PLACEHOLDER"
            ErrorCataLog(0, 0023) = 1000: ErrorCataLog(1, 0023) = "PLACEHOLDER"
            ErrorCataLog(0, 0024) = 1000: ErrorCataLog(1, 0024) = "PLACEHOLDER"
            ErrorCataLog(0, 0025) = 1000: ErrorCataLog(1, 0025) = "PLACEHOLDER"
            ErrorCataLog(0, 0026) = 1000: ErrorCataLog(1, 0026) = "PLACEHOLDER"
            ErrorCataLog(0, 0027) = 1000: ErrorCataLog(1, 0027) = "PLACEHOLDER"
            ErrorCataLog(0, 0028) = 1000: ErrorCataLog(1, 0028) = "PLACEHOLDER"
            ErrorCataLog(0, 0029) = 1000: ErrorCataLog(1, 0029) = "PLACEHOLDER"
            ErrorCataLog(0, 0030) = 1000: ErrorCataLog(1, 0030) = "PLACEHOLDER"
            ErrorCataLog(0, 0031) = 1000: ErrorCataLog(1, 0031) = "PLACEHOLDER"
            ErrorCataLog(0, 0032) = 1000: ErrorCataLog(1, 0032) = "PLACEHOLDER"
            ErrorCataLog(0, 0033) = 1000: ErrorCataLog(1, 0033) = "PLACEHOLDER"
            ErrorCataLog(0, 0034) = 1000: ErrorCataLog(1, 0034) = "PLACEHOLDER"
            ErrorCataLog(0, 0035) = 1000: ErrorCataLog(1, 0035) = "PLACEHOLDER"
            ErrorCataLog(0, 0036) = 1000: ErrorCataLog(1, 0036) = "PLACEHOLDER"
            ErrorCataLog(0, 0037) = 1000: ErrorCataLog(1, 0037) = "PLACEHOLDER"
            ErrorCataLog(0, 0038) = 1000: ErrorCataLog(1, 0038) = "PLACEHOLDER"
            ErrorCataLog(0, 0039) = 1000: ErrorCataLog(1, 0039) = "PLACEHOLDER"
            ErrorCataLog(0, 0040) = 1000: ErrorCataLog(1, 0040) = "PLACEHOLDER"
            ErrorCataLog(0, 0041) = 1000: ErrorCataLog(1, 0041) = "PLACEHOLDER"
            ErrorCataLog(0, 0042) = 1000: ErrorCataLog(1, 0042) = "PLACEHOLDER"
            ErrorCataLog(0, 0043) = 1000: ErrorCataLog(1, 0043) = "PLACEHOLDER"
            ErrorCataLog(0, 0044) = 1000: ErrorCataLog(1, 0044) = "PLACEHOLDER"
            ErrorCataLog(0, 0045) = 1000: ErrorCataLog(1, 0045) = "PLACEHOLDER"
            ErrorCataLog(0, 0046) = 1000: ErrorCataLog(1, 0046) = "PLACEHOLDER"
            ErrorCataLog(0, 0047) = 1000: ErrorCataLog(1, 0047) = "PLACEHOLDER"
            ErrorCataLog(0, 0048) = 1000: ErrorCataLog(1, 0048) = "PLACEHOLDER"
            ErrorCataLog(0, 0049) = 1000: ErrorCataLog(1, 0049) = "PLACEHOLDER"
            ErrorCataLog(0, 0050) = 1000: ErrorCataLog(1, 0050) = "PLACEHOLDER"
            ErrorCataLog(0, 0051) = 1000: ErrorCataLog(1, 0051) = "PLACEHOLDER"
            ErrorCataLog(0, 0052) = 1000: ErrorCataLog(1, 0052) = "PLACEHOLDER"
            ErrorCataLog(0, 0053) = 1000: ErrorCataLog(1, 0053) = "PLACEHOLDER"
            ErrorCataLog(0, 0054) = 1000: ErrorCataLog(1, 0054) = "PLACEHOLDER"
            ErrorCataLog(0, 0055) = 1000: ErrorCataLog(1, 0055) = "PLACEHOLDER"
            ErrorCataLog(0, 0056) = 1000: ErrorCataLog(1, 0056) = "PLACEHOLDER"
            ErrorCataLog(0, 0057) = 1000: ErrorCataLog(1, 0057) = "PLACEHOLDER"
            ErrorCataLog(0, 0058) = 1000: ErrorCataLog(1, 0058) = "PLACEHOLDER"
            ErrorCataLog(0, 0059) = 1000: ErrorCataLog(1, 0059) = "PLACEHOLDER"
            ErrorCataLog(0, 0060) = 1000: ErrorCataLog(1, 0060) = "PLACEHOLDER"
            ErrorCataLog(0, 0061) = 1000: ErrorCataLog(1, 0061) = "PLACEHOLDER"
            ErrorCataLog(0, 0062) = 1000: ErrorCataLog(1, 0062) = "PLACEHOLDER"
            ErrorCataLog(0, 0063) = 1000: ErrorCataLog(1, 0063) = "PLACEHOLDER"
            ErrorCataLog(0, 0064) = 1000: ErrorCataLog(1, 0064) = "PLACEHOLDER"
            ErrorCataLog(0, 0065) = 1000: ErrorCataLog(1, 0065) = "PLACEHOLDER"
            ErrorCataLog(0, 0066) = 1000: ErrorCataLog(1, 0066) = "PLACEHOLDER"
            ErrorCataLog(0, 0067) = 1000: ErrorCataLog(1, 0067) = "PLACEHOLDER"
            ErrorCataLog(0, 0068) = 1000: ErrorCataLog(1, 0068) = "PLACEHOLDER"
            ErrorCataLog(0, 0069) = 1000: ErrorCataLog(1, 0069) = "PLACEHOLDER"
            ErrorCataLog(0, 0070) = 1000: ErrorCataLog(1, 0070) = "PLACEHOLDER"
            ErrorCataLog(0, 0071) = 1000: ErrorCataLog(1, 0071) = "PLACEHOLDER"
            ErrorCataLog(0, 0072) = 1000: ErrorCataLog(1, 0072) = "PLACEHOLDER"
            ErrorCataLog(0, 0073) = 1000: ErrorCataLog(1, 0073) = "PLACEHOLDER"
            ErrorCataLog(0, 0074) = 1000: ErrorCataLog(1, 0074) = "PLACEHOLDER"
            ErrorCataLog(0, 0075) = 1000: ErrorCataLog(1, 0075) = "PLACEHOLDER"
            ErrorCataLog(0, 0076) = 1000: ErrorCataLog(1, 0076) = "PLACEHOLDER"
            ErrorCataLog(0, 0077) = 1000: ErrorCataLog(1, 0077) = "PLACEHOLDER"
            ErrorCataLog(0, 0078) = 1000: ErrorCataLog(1, 0078) = "PLACEHOLDER"
            ErrorCataLog(0, 0079) = 1000: ErrorCataLog(1, 0079) = "PLACEHOLDER"
            ErrorCataLog(0, 0080) = 1000: ErrorCataLog(1, 0080) = "PLACEHOLDER"
            ErrorCataLog(0, 0081) = 1000: ErrorCataLog(1, 0081) = "PLACEHOLDER"
            ErrorCataLog(0, 0082) = 1000: ErrorCataLog(1, 0082) = "PLACEHOLDER"
            ErrorCataLog(0, 0083) = 1000: ErrorCataLog(1, 0083) = "PLACEHOLDER"
            ErrorCataLog(0, 0084) = 1000: ErrorCataLog(1, 0084) = "PLACEHOLDER"
            ErrorCataLog(0, 0085) = 1000: ErrorCataLog(1, 0085) = "PLACEHOLDER"
            ErrorCataLog(0, 0086) = 1000: ErrorCataLog(1, 0086) = "PLACEHOLDER"
            ErrorCataLog(0, 0087) = 1000: ErrorCataLog(1, 0087) = "PLACEHOLDER"
            ErrorCataLog(0, 0088) = 1000: ErrorCataLog(1, 0088) = "PLACEHOLDER"
            ErrorCataLog(0, 0089) = 1000: ErrorCataLog(1, 0089) = "PLACEHOLDER"
            ErrorCataLog(0, 0090) = 1000: ErrorCataLog(1, 0090) = "PLACEHOLDER"
            ErrorCataLog(0, 0091) = 1000: ErrorCataLog(1, 0091) = "PLACEHOLDER"
            ErrorCataLog(0, 0092) = 1000: ErrorCataLog(1, 0092) = "PLACEHOLDER"
            ErrorCataLog(0, 0093) = 1000: ErrorCataLog(1, 0093) = "PLACEHOLDER"
            ErrorCataLog(0, 0094) = 1000: ErrorCataLog(1, 0094) = "PLACEHOLDER"
            ErrorCataLog(0, 0095) = 1000: ErrorCataLog(1, 0095) = "PLACEHOLDER"
            ErrorCataLog(0, 0096) = 1000: ErrorCataLog(1, 0096) = "PLACEHOLDER"
            ErrorCataLog(0, 0097) = 1000: ErrorCataLog(1, 0097) = "PLACEHOLDER"
            ErrorCataLog(0, 0098) = 1000: ErrorCataLog(1, 0098) = "PLACEHOLDER"
            ErrorCataLog(0, 0099) = 1000: ErrorCataLog(1, 0099) = "PLACEHOLDER"
            Initialized = True
        End If
    End Function
'
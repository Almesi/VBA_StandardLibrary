VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private       ErrorCatalog(1, 99)   As Variant
Private       Initialized           As Boolean

Private Const p_IS_ERROR            As Boolean = True
Private Const EMPTY_ERROR           As Variant = Empty
Private Const SEVERITY_BREAK        As Long    = 1000
Private Const ERROR_QUESTION        As Long    = 0001
Private Const ERROR_CATEGORY        As Long    = 0002

Private       p_ShowError           As Boolean
Private       p_LogError            As Boolean
Private       p_LoggingDestination  As Variant
Private       p_ShowDestination     As Variant

Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long


' Public Errormethods
    Public Property Get IS_ERROR()
        IS_ERROR = p_IS_ERROR
    End Property

    Public Property Let ShowError(Value As Boolean)
        p_ShowError = Value
    End Property
    Public Property Let LogError(Value As Boolean)
        p_LogError = Value
    End Property

    Public Property Set LoggingDestination(n_Destination As Variant)
        If IsObject(n_Destination) Then
            Set p_LoggingDestination = n_Destination
        Else
            p_LoggingDestination = n_Destination
        End If
    End Property
    Public Property Set ShowDestination(n_Destination As Variant)
        If IsObject(n_Destination) Then
            Set p_ShowDestination = n_Destination
        Else
            p_ShowDestination = n_Destination
        End If
    End Property


    Public Function Create(Optional n_ShowError As Boolean = True, Optional n_LogError As Boolean = True, Optional n_LoggingDestination As Variant, Optional n_ShowDestination As Variant) As std_Error
        Set Create = New std_Error
        Let Create.ShowError          = n_ShowError
        Let Create.LogError           = n_LogError
        Set Create.LoggingDestination = n_LoggingDestination
        Set Create.ShowDestination    = n_ShowDestination

    End Function

    Public Function Handle(Catalog() As Variant, Index As Long, ParamArray ErrorValues()) As Boolean

        Dim Severity As Integer
        Dim Arguments As Variant

        Call ValidateCatalog(Catalog)
        Arguments = ErrorValues
        Severity = Catalog(0, Index)
        If Severity = ERROR_QUESTION Then
            Handle = Ask(Catalog, Index, Arguments)
        Else
            Call Showing(Catalog, Index, Arguments)
            Handle = IS_ERROR
        End If
        Call Logging(Catalog, Index, Arguments) 
        If Severity > SEVERITY_BREAK Then End

    End Function
'



' Private Errormethods
    Private Sub Logging(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty)

        If p_LogError Then
            Dim i As Long
            Dim fso As Object
            Dim TextStream As Object
            Dim Text As String
            Text = GetMessage(Catalog, Index, ErrorValues)
            Set fso = CreateObject("Scripting.FileSystemObject")
            Select Case True
                Case (p_LoggingDestination Is Nothing)         : Debug.Print Text
                Case TypeName(p_LoggingDestination) = "Console": Call Console.PrintEnter(Text, (&H100000FF - &H10000000))
                Case TypeName(p_LoggingDestination) = "Range"
                    Do Until p_LoggingDestination.Offset(i, 0).Formula = ""
                        i = i + 1
                    Loop
                    p_LoggingDestination.Offset(i, 0).Formula = Text
                Case TypeName(p_LoggingDestination) = "File", TypeName(p_LoggingDestination) = "String"
                    If TypeName(p_LoggingDestination) = "String" Then
                        If fso.FileExists(p_LoggingDestination) Then
                        Set TextStream = fso.GetFile(p_LoggingDestination).OpenAsTextStream(2, 2)
                        End If
                    Else
                        Set TextStream = p_LoggingDestination.OpenAsTextStream(2, 2) '(ForWriting, TristateUseDefault)
                    End If
                    TextStream.WriteLine Text
                    TextStream.Close
            End Select
        End If
    End Sub
    
    Private Sub Showing(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty)
        If p_ShowError Then
            Dim Temp As Variant
            Select Case True
                Case (p_ShowDestination Is Nothing): Temp = MsgBox(GetMessage(Catalog, Index, ErrorValues), vbExclamation, "ERROR")
            End Select
        End If
    End Sub
    
    Private Function Ask(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty) As Boolean
        If p_ShowError Then
            Dim Temp As Variant
            Dim ArrayAnswer() As Variant
            Dim ArrayMessage() As Variant
            ArrayAnswer = Array("y","n")
            ArrayMessage = Array("Answer is Yes", "Answer is No")
            Select Case True
                Case (p_ShowDestination Is Nothing)            : Temp = MsgBox(GetMessage(Catalog, Index, ErrorValues), vbExclamation, "ERROR")
                Case TypeName(p_LoggingDestination) = "Console": Call Console.CheckInput(GetMessage(Catalog, Index, ErrorValues), ArrayAnswer, ArrayMessage)
            End Select
            If Temp = False Then Ask = IS_ERROR
        End If
    End Function

    Private Function ValidateCatalog(Catalog() As Variant) As Boolean

        Dim ErrorCode As Long
        Dim ErrorValue(1) As Long
        On Error GoTo Error
        ErrorCode = 3
        If IsArray(Catalog) = False Then GoTo Error
        ErrorCode = 1
        If Ubound(Catalog, 1) <> 1 Or Ubound(Catalog, 2) =< 0 Then GoTo Error
        ErrorCode = 2
        If Catalog(0,0) <> ERROR_CATEGORY Then
            ErrorValue(0) = Catalog(0,0)
            ErrorValue(1) = Catalog(1,0)
        End If
        Exit Function

        Error:
        Call Handle(ErrorCatalog, ErrorCode, ErrorValue)
        ValidateCatalog = IS_ERROR

    End Function
    
    Private Function GetMessage(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty) As String
    
        Dim ErrorMessage As String
        Dim String1 As String
        Dim i As Long
        String1 = Catalog(1, Index)
        ErrorMessage = "Category: " & Catalog(1, 0)                              & vbCrLf & _
                       "Severity: " & Catalog(0, Index)                          & vbCrLf & _
                       "Index   : " & Index                                      & vbCrLf & _
                       "Message : " & String1                                    & vbCrLf
        If ISArray(ErrorValues) Then
            For i = 0 To UBound(ErrorValues)
                ErrorMessage = ErrorMessage & "Value" & i & "  : " & ErrorValues(i) & vbCrLf
            Next
        Else
                ErrorMessage = ErrorMessage & "Value" & i & "  : " & ErrorValues & vbCrLf
        End If
        ErrorMessage = ErrorMessage & "------------------------------------------------------------------------------"
        GetMessage = ErrorMessage
    
    End Function

    Private Sub Class_Initialize()
        ProtInit
    End Sub

    Private Sub Class_Terminate()
    End Sub
    
    Private Sub ProtInit()
        If Initialized = False Then
            ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "std_Error"
            ErrorCatalog(0, 0001) = 1001: ErrorCatalog(1, 0001) = "Error-Catalog not a 2D Array with size (1,X)"
            ErrorCatalog(0, 0002) = 1001: ErrorCatalog(1, 0002) = "Error-Catalog not properly initialize"
            ErrorCatalog(0, 0003) = 1001: ErrorCatalog(1, 0003) = "Error-Catalog not an Array"
            ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "PLACEHOLDER"
            ErrorCatalog(0, 0005) = 1000: ErrorCatalog(1, 0005) = "PLACEHOLDER"
            ErrorCatalog(0, 0006) = 1000: ErrorCatalog(1, 0006) = "PLACEHOLDER"
            ErrorCatalog(0, 0007) = 1000: ErrorCatalog(1, 0007) = "PLACEHOLDER"
            ErrorCatalog(0, 0008) = 1000: ErrorCatalog(1, 0008) = "PLACEHOLDER"
            ErrorCatalog(0, 0009) = 1000: ErrorCatalog(1, 0009) = "PLACEHOLDER"
            ErrorCatalog(0, 0010) = 1000: ErrorCatalog(1, 0010) = "PLACEHOLDER"
            ErrorCatalog(0, 0011) = 1000: ErrorCatalog(1, 0011) = "PLACEHOLDER"
            ErrorCatalog(0, 0012) = 1000: ErrorCatalog(1, 0012) = "PLACEHOLDER"
            ErrorCatalog(0, 0013) = 1000: ErrorCatalog(1, 0013) = "PLACEHOLDER"
            ErrorCatalog(0, 0014) = 1000: ErrorCatalog(1, 0014) = "PLACEHOLDER"
            ErrorCatalog(0, 0015) = 1000: ErrorCatalog(1, 0015) = "PLACEHOLDER"
            ErrorCatalog(0, 0016) = 1000: ErrorCatalog(1, 0016) = "PLACEHOLDER"
            ErrorCatalog(0, 0017) = 1000: ErrorCatalog(1, 0017) = "PLACEHOLDER"
            ErrorCatalog(0, 0018) = 1000: ErrorCatalog(1, 0018) = "PLACEHOLDER"
            ErrorCatalog(0, 0019) = 1000: ErrorCatalog(1, 0019) = "PLACEHOLDER"
            ErrorCatalog(0, 0020) = 1000: ErrorCatalog(1, 0020) = "PLACEHOLDER"
            ErrorCatalog(0, 0021) = 1000: ErrorCatalog(1, 0021) = "PLACEHOLDER"
            ErrorCatalog(0, 0022) = 1000: ErrorCatalog(1, 0022) = "PLACEHOLDER"
            ErrorCatalog(0, 0023) = 1000: ErrorCatalog(1, 0023) = "PLACEHOLDER"
            ErrorCatalog(0, 0024) = 1000: ErrorCatalog(1, 0024) = "PLACEHOLDER"
            ErrorCatalog(0, 0025) = 1000: ErrorCatalog(1, 0025) = "PLACEHOLDER"
            ErrorCatalog(0, 0026) = 1000: ErrorCatalog(1, 0026) = "PLACEHOLDER"
            ErrorCatalog(0, 0027) = 1000: ErrorCatalog(1, 0027) = "PLACEHOLDER"
            ErrorCatalog(0, 0028) = 1000: ErrorCatalog(1, 0028) = "PLACEHOLDER"
            ErrorCatalog(0, 0029) = 1000: ErrorCatalog(1, 0029) = "PLACEHOLDER"
            ErrorCatalog(0, 0030) = 1000: ErrorCatalog(1, 0030) = "PLACEHOLDER"
            ErrorCatalog(0, 0031) = 1000: ErrorCatalog(1, 0031) = "PLACEHOLDER"
            ErrorCatalog(0, 0032) = 1000: ErrorCatalog(1, 0032) = "PLACEHOLDER"
            ErrorCatalog(0, 0033) = 1000: ErrorCatalog(1, 0033) = "PLACEHOLDER"
            ErrorCatalog(0, 0034) = 1000: ErrorCatalog(1, 0034) = "PLACEHOLDER"
            ErrorCatalog(0, 0035) = 1000: ErrorCatalog(1, 0035) = "PLACEHOLDER"
            ErrorCatalog(0, 0036) = 1000: ErrorCatalog(1, 0036) = "PLACEHOLDER"
            ErrorCatalog(0, 0037) = 1000: ErrorCatalog(1, 0037) = "PLACEHOLDER"
            ErrorCatalog(0, 0038) = 1000: ErrorCatalog(1, 0038) = "PLACEHOLDER"
            ErrorCatalog(0, 0039) = 1000: ErrorCatalog(1, 0039) = "PLACEHOLDER"
            ErrorCatalog(0, 0040) = 1000: ErrorCatalog(1, 0040) = "PLACEHOLDER"
            ErrorCatalog(0, 0041) = 1000: ErrorCatalog(1, 0041) = "PLACEHOLDER"
            ErrorCatalog(0, 0042) = 1000: ErrorCatalog(1, 0042) = "PLACEHOLDER"
            ErrorCatalog(0, 0043) = 1000: ErrorCatalog(1, 0043) = "PLACEHOLDER"
            ErrorCatalog(0, 0044) = 1000: ErrorCatalog(1, 0044) = "PLACEHOLDER"
            ErrorCatalog(0, 0045) = 1000: ErrorCatalog(1, 0045) = "PLACEHOLDER"
            ErrorCatalog(0, 0046) = 1000: ErrorCatalog(1, 0046) = "PLACEHOLDER"
            ErrorCatalog(0, 0047) = 1000: ErrorCatalog(1, 0047) = "PLACEHOLDER"
            ErrorCatalog(0, 0048) = 1000: ErrorCatalog(1, 0048) = "PLACEHOLDER"
            ErrorCatalog(0, 0049) = 1000: ErrorCatalog(1, 0049) = "PLACEHOLDER"
            ErrorCatalog(0, 0050) = 1000: ErrorCatalog(1, 0050) = "PLACEHOLDER"
            ErrorCatalog(0, 0051) = 1000: ErrorCatalog(1, 0051) = "PLACEHOLDER"
            ErrorCatalog(0, 0052) = 1000: ErrorCatalog(1, 0052) = "PLACEHOLDER"
            ErrorCatalog(0, 0053) = 1000: ErrorCatalog(1, 0053) = "PLACEHOLDER"
            ErrorCatalog(0, 0054) = 1000: ErrorCatalog(1, 0054) = "PLACEHOLDER"
            ErrorCatalog(0, 0055) = 1000: ErrorCatalog(1, 0055) = "PLACEHOLDER"
            ErrorCatalog(0, 0056) = 1000: ErrorCatalog(1, 0056) = "PLACEHOLDER"
            ErrorCatalog(0, 0057) = 1000: ErrorCatalog(1, 0057) = "PLACEHOLDER"
            ErrorCatalog(0, 0058) = 1000: ErrorCatalog(1, 0058) = "PLACEHOLDER"
            ErrorCatalog(0, 0059) = 1000: ErrorCatalog(1, 0059) = "PLACEHOLDER"
            ErrorCatalog(0, 0060) = 1000: ErrorCatalog(1, 0060) = "PLACEHOLDER"
            ErrorCatalog(0, 0061) = 1000: ErrorCatalog(1, 0061) = "PLACEHOLDER"
            ErrorCatalog(0, 0062) = 1000: ErrorCatalog(1, 0062) = "PLACEHOLDER"
            ErrorCatalog(0, 0063) = 1000: ErrorCatalog(1, 0063) = "PLACEHOLDER"
            ErrorCatalog(0, 0064) = 1000: ErrorCatalog(1, 0064) = "PLACEHOLDER"
            ErrorCatalog(0, 0065) = 1000: ErrorCatalog(1, 0065) = "PLACEHOLDER"
            ErrorCatalog(0, 0066) = 1000: ErrorCatalog(1, 0066) = "PLACEHOLDER"
            ErrorCatalog(0, 0067) = 1000: ErrorCatalog(1, 0067) = "PLACEHOLDER"
            ErrorCatalog(0, 0068) = 1000: ErrorCatalog(1, 0068) = "PLACEHOLDER"
            ErrorCatalog(0, 0069) = 1000: ErrorCatalog(1, 0069) = "PLACEHOLDER"
            ErrorCatalog(0, 0070) = 1000: ErrorCatalog(1, 0070) = "PLACEHOLDER"
            ErrorCatalog(0, 0071) = 1000: ErrorCatalog(1, 0071) = "PLACEHOLDER"
            ErrorCatalog(0, 0072) = 1000: ErrorCatalog(1, 0072) = "PLACEHOLDER"
            ErrorCatalog(0, 0073) = 1000: ErrorCatalog(1, 0073) = "PLACEHOLDER"
            ErrorCatalog(0, 0074) = 1000: ErrorCatalog(1, 0074) = "PLACEHOLDER"
            ErrorCatalog(0, 0075) = 1000: ErrorCatalog(1, 0075) = "PLACEHOLDER"
            ErrorCatalog(0, 0076) = 1000: ErrorCatalog(1, 0076) = "PLACEHOLDER"
            ErrorCatalog(0, 0077) = 1000: ErrorCatalog(1, 0077) = "PLACEHOLDER"
            ErrorCatalog(0, 0078) = 1000: ErrorCatalog(1, 0078) = "PLACEHOLDER"
            ErrorCatalog(0, 0079) = 1000: ErrorCatalog(1, 0079) = "PLACEHOLDER"
            ErrorCatalog(0, 0080) = 1000: ErrorCatalog(1, 0080) = "PLACEHOLDER"
            ErrorCatalog(0, 0081) = 1000: ErrorCatalog(1, 0081) = "PLACEHOLDER"
            ErrorCatalog(0, 0082) = 1000: ErrorCatalog(1, 0082) = "PLACEHOLDER"
            ErrorCatalog(0, 0083) = 1000: ErrorCatalog(1, 0083) = "PLACEHOLDER"
            ErrorCatalog(0, 0084) = 1000: ErrorCatalog(1, 0084) = "PLACEHOLDER"
            ErrorCatalog(0, 0085) = 1000: ErrorCatalog(1, 0085) = "PLACEHOLDER"
            ErrorCatalog(0, 0086) = 1000: ErrorCatalog(1, 0086) = "PLACEHOLDER"
            ErrorCatalog(0, 0087) = 1000: ErrorCatalog(1, 0087) = "PLACEHOLDER"
            ErrorCatalog(0, 0088) = 1000: ErrorCatalog(1, 0088) = "PLACEHOLDER"
            ErrorCatalog(0, 0089) = 1000: ErrorCatalog(1, 0089) = "PLACEHOLDER"
            ErrorCatalog(0, 0090) = 1000: ErrorCatalog(1, 0090) = "PLACEHOLDER"
            ErrorCatalog(0, 0091) = 1000: ErrorCatalog(1, 0091) = "PLACEHOLDER"
            ErrorCatalog(0, 0092) = 1000: ErrorCatalog(1, 0092) = "PLACEHOLDER"
            ErrorCatalog(0, 0093) = 1000: ErrorCatalog(1, 0093) = "PLACEHOLDER"
            ErrorCatalog(0, 0094) = 1000: ErrorCatalog(1, 0094) = "PLACEHOLDER"
            ErrorCatalog(0, 0095) = 1000: ErrorCatalog(1, 0095) = "PLACEHOLDER"
            ErrorCatalog(0, 0096) = 1000: ErrorCatalog(1, 0096) = "PLACEHOLDER"
            ErrorCatalog(0, 0097) = 1000: ErrorCatalog(1, 0097) = "PLACEHOLDER"
            ErrorCatalog(0, 0098) = 1000: ErrorCatalog(1, 0098) = "PLACEHOLDER"
            ErrorCatalog(0, 0099) = 1000: ErrorCatalog(1, 0099) = "PLACEHOLDER"
            Initialized = True
        End If
    End Sub
'
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

' Array of all Errormessages
    ' First  Dimension is Severity/Message
    ' Second Dimension is Index
Private ErrorCatalog(1, 99) As Variant
' Used once to initialize all Errormessages
Private Initialized As Boolean

' Used to determine, if it is an Error
Private Const p_IS_ERROR As Boolean = True

' What to show of the error
Private p_ShowError As Boolean
Private p_PrintError As Boolean

' Stops process if this Value is lower than the Error severity
Private Const SEVERITY_BREAK As Long = 1000

' Runs a Question and handles Error accordingly. Should be below SEVERITY_BREAK
Private Const ERROR_QUESTION As Long = 0001

' Determines this Errormessage as the category-name, dont change this
Private Const ERROR_CATEGORY As Long = 0002

' Standard Value if no ErrorValue is passed
Private Const EMPTY_ERROR As Variant = Empty

' Pass your goal to log here. For example a Range or textfile
Private LogDestination As Variant

' De-/-activate LogMode
Private LogMode As Long
Private Enum LogModeEnum
    Standard = 0    
    Console = 1
    Sheet = 2
End Enum

Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long


' Public Errormethods
    Public Property Get IS_ERROR()
        IS_ERROR = p_IS_ERROR
    End Property

    Public Property Let ShowError(Value As Boolean)
        p_ShowError = Value
    End Property
    Public Property Let PrintError(Value As Boolean)
        p_PrintError = Value
    End Property

    Public Property Let Mode(Value As Long)
        LogMode = Value
    End Property
    Public Property Get Mode() As Long
        Mode = LogMode
    End Property

    Public Property Set Destination(n_Destination As Variant)
        Set LogDestination = n_Destination
    End Property


    ' Handles Errormessage
    Public Function Handle(Catalog() As Variant, Index As Long, ParamArray ErrorValues()) As Boolean

        Dim Severity As Integer
        Dim Arguments As Variant

        Arguments = ErrorValues
        If IsArray(Catalog) = False Then
            Showw ErrorCatalog, 35
            Printt ErrorCatalog, 35
            Handle = IS_ERROR
            End
            Exit Function
        End If
        On Error Resume Next
        If Ubound(Catalog, 1) <> 1 Or Ubound(Catalog, 2) =< 0 Then
            Showw ErrorCatalog, 33
            Printt ErrorCatalog, 33
            Handle = IS_ERROR
            End
            Exit Function
        End If
            If Catalog(0,0) <> ERROR_CATEGORY Then
            Dim Temp(1)
            Temp(0) = Catalog(0,0)
            Temp(1) = Catalog(1,0)
            Showw ErrorCatalog, 34, Temp
            Printt ErrorCatalog, 34, Temp
            Handle = IS_ERROR
            End
            Exit Function
        Else
        End If
        Severity = Catalog(0, Index)
        If Severity = ERROR_QUESTION Then
            Handle = Ask(Catalog, Index, Arguments)
        Else
            Showw Catalog, Index, Arguments
            Handle = IS_ERROR
        End If
        Printt Catalog, Index, Arguments 
        If Severity > SEVERITY_BREAK Then End

    End Function

    ' Compare Numbers or Text and Handle Errors
    Public Function Variable(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean

        If Operator <> Empty Then
            Select Case UCase(Operator)
                Case "=", "IS"         : If FirstValue =  SecondValue Then Variable = Handle(ErrorCatalog, 16, FirstValue, SecondValue): Exit Function
                Case "<>", "NOT", "!=" : If FirstValue <> SecondValue Then Variable = Handle(ErrorCatalog, 07, FirstValue, SecondValue): Exit Function
                Case "<"               : If FirstValue <  SecondValue Then Variable = Handle(ErrorCatalog, 10, FirstValue, SecondValue): Exit Function
                Case ">"               : If FirstValue >  SecondValue Then Variable = Handle(ErrorCatalog, 11, FirstValue, SecondValue): Exit Function
                Case "=<", "<="        : If FirstValue =< SecondValue Then Variable = Handle(ErrorCatalog, 08, FirstValue, SecondValue): Exit Function
                Case ">=", "=>"        : If FirstValue >= SecondValue Then Variable = Handle(ErrorCatalog, 09, FirstValue, SecondValue): Exit Function
                Case Else
            End Select
        Else
            If FirstValue = Empty                                     Then Variable = Handle(ErrorCatalog, 03, "FirstValue"):            Exit Function
        End If
        If MinValue <> Empty Then         If FirstValue < MinValue    Then Variable = Handle(ErrorCatalog, 10, FirstValue, MinValue):    Exit Function
        If MaxValue <> Empty Then         If FirstValue > MaxValue    Then Variable = Handle(ErrorCatalog, 11, FirstValue, MaxValue):    Exit Function
        

    End Function

    ' Compare Numbers or Text and Handle Errors
    Public Function Object(FirstValue As Object, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Object = EMPTY_ERROR) As Boolean

        If Operator <> Empty Then
            Select Case UCase(Operator)
                Case "=", "IS"        : If FirstValue     IS SecondValue Then Object = Handle(ErrorCatalog, 16, "FirstValue", "SecondValue"): Exit Function
                Case "<>", "NOT", "!=": If Not FirstValue IS SecondValue Then Object = Handle(ErrorCatalog, 07, "FirstValue", "SecondValue"): Exit Function
            End Select
        End If
        If FirstValue Is Nothing                                         Then Object = Handle(ErrorCatalog, 04, "FirstValue"):                Exit Function

    End Function

    ' Compare Strings and Handle Errors
    Public Function Strings(Text As String, Operator As String, SecondText As String) As Boolean

        If Operator <> Empty Then
            Select Case UCase(Operator)
                Case "=", "IS"           : If Text =        SecondText Then Strings = Handle(ErrorCatalog, 16, Text, SecondText): Exit Function
                Case "<>", "NOT", "!="   : If Text <>       SecondText Then Strings = Handle(ErrorCatalog, 07, Text, SecondText): Exit Function
                Case "<"                 : If Text <        SecondText Then Strings = Handle(ErrorCatalog, 11, Text, SecondText): Exit Function
                Case ">"                 : If Text >        SecondText Then Strings = Handle(ErrorCatalog, 10, Text, SecondText): Exit Function
                Case "=<", "<="          : If Text =<       SecondText Then Strings = Handle(ErrorCatalog, 09, Text, SecondText): Exit Function
                Case ">=", "=>"          : If Text >=       SecondText Then Strings = Handle(ErrorCatalog, 08, Text, SecondText): Exit Function
                Case "LIKE"              : If Text Like     SecondText Then Strings = Handle(ErrorCatalog, 18, Text, SecondText): Exit Function
                Case "NOT LIKE", "UNLIKE": If Not Text Like SecondText Then Strings = Handle(ErrorCatalog, 17, Text, SecondText): Exit Function
                Case Else
            End Select
        End If
        If Text = Empty                                                Then Strings = Handle(ErrorCatalog, 03, "Text"):           Exit Function

    End Function

    ' Superset of Variable, used to check if it is a number and Handle Errors
    Public Function Number(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean
        
        If IsNumeric(FirstValue) = False Then
            Number = Handle(ErrorCatalog, 19, FirstValue)
            Exit Function
        Else
            Number = Variable(FirstValue, Operator, SecondValue, MinValue, MaxValue)
        End If

    End Function

    ' Superset of Variable, used to check if it is a date and Handle Errors
    Public Function Dates(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean
        
        If IsDate(FirstValue) = False Then
            Dates = Handle(ErrorCatalog, 19, Text, SecondText)
            Exit Function
        Else
            Dates = Variable(FirstValue, Operator, SecondValue, MinValue, MaxValue)
        End If

    End Function

    ' Handles File Validation and Handle Errors
    Public Function File(FilePath As String, Optional ShouldExist As Boolean = True) As Boolean

        If ShouldExist = True Then
            If Dir(FilePath) = "" Then  File = Handle(ErrorCatalog, 20, FilePath)
        Else
            If Dir(FilePath) <> "" Then File = Handle(ErrorCatalog, 21, FilePath)
        End If   

    End Function

    ' Check Connection to specified Computer and Handle Errors
    Public Function Connection(Optional Computer As String = ".", Optional ShouldExist As Boolean = True) As Boolean

        Dim objWMIService As Object
        Dim colItems As Object
        Dim objItem As Object

        Set objWMIService = GetObject("winmgmts:\\" & Computer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
        If ShouldExist = True Then
            If colItems.Count = 0 Then
                Connection = Handle(ErrorCatalog, 22, Computer)
            Else
                For Each objItem In colItems
                    If objItem.IPAddress(0) <> "" Then Exit Function
                Next
                Connection = Handle(ErrorCatalog, 23, Computer)
            End If
        Else
            If colItems.Count <> 0 Then
                Connection = Handle(ErrorCatalog, 25, Computer)
            Else
                For Each objItem In colItems
                    If objItem.IPAddress(0) = "" Then Exit Function
                Next
                Connection = Handle(ErrorCatalog, 26, Computer)
            End If
        End If

    End Function

        ' INTERNET_CONNECTION_MODEM   = &H1
        ' INTERNET_CONNECTION_LAN     = &H2
        ' INTERNET_CONNECTION_PROXY   = &H4
        ' INTERNET_CONNECTION_OFFLINE = &H20
    Public Function InternetConnection(Optional ConnectType As Long = 0) As Boolean

        Dim L As Long
        Dim R As Long
        R = InternetGetConnectedState(L, 0&)
        Select Case R
            Case &H00:                         InternetConnection = Handle(ErrorCatalog, 35, ConnectType)
            Case &H01: If ConnectType = R Then InternetConnection = Handle(ErrorCatalog, 36, ConnectType)
            Case &H02: If ConnectType = R Then InternetConnection = Handle(ErrorCatalog, 37, ConnectType)
            Case &H04: If ConnectType = R Then InternetConnection = Handle(ErrorCatalog, 38, ConnectType)
            Case &H20: If ConnectType = R Then InternetConnection = Handle(ErrorCatalog, 39, ConnectType)
            Case Else
        End Select
            
    End Function
                
            

    ' Check DatabaseConnection and Handle Errors
    Public Function ConnectToDatabase(DataBasePath As String) As Boolean

        On Error GoTo ErrorHandler
        Dim Conn As Object
        Set Conn = CreateObject("ADODB.Connection")
        Conn.Open DataBasePath
        If Conn.State = 1 Then
            ConnectToDatabase = True
        Else
            ConnectToDatabase = Handle(ErrorCatalog, 27, DataBasePath)
        End If
        Exit Function
        ErrorHandler:
        ConnectToDatabase = Handle(ErrorCatalog, 24, "ConnectToDatabase", DataBasePath)

    End Function

    ' Check if passed Variable is InputType according to ShouldBe
    Public Function DataType(Value As Variant, InputType As String, Optional ShouldBe As Boolean = True) As Boolean

        Dim Inputt As Boolean

        Select Case UCase(InputType)
            Case "VARIANT":         If VarType(Value)                                                 = vbVariant  Then Inputt = True
            Case "STRING":          If VarType(Value)                                                 = vbString   Then Inputt = True
            Case "NUMBER":          If IsNumeric(Value)                                               = True       Then Inputt = True
            Case "DATE":            If IsDate(Value)                                                  = True       Then Inputt = True
            Case "BOOLEAN":         If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbBoolean  Then Inputt = True
            Case "INTEGER", "INT":  If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbInteger  Then Inputt = True
            Case "LONG":            If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbLong     Then Inputt = True
            Case "LONGLONG":        If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbLongLong Then Inputt = True
            Case "DOUBLE":          If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbDouble   Then Inputt = True
            Case "SINGLE", "FLOAT": If IsNumeric(Value) Then Value = CDbl(Value): If VarType(Value)   = vbSingle   Then Inputt = True
            Case "EMPTY":           If IsEmpty(Value)                                                 = True       Then Inputt = True
            Case "NULL":            If IsNull(Value)                                                  = True       Then Inputt = True
            Case Else:             DataType = Handle(ErrorCatalog, 30, Value, InputType)
        End Select
        If Not Inputt Xor True Then
            Exit Function
        Else
            If Inputt = True Then
                DataType = Handle(ErrorCatalog, 28, Value, InputType)
            Else
                DataType = Handle(ErrorCatalog, 29, Value, InputType)
            End If
        End If        
        Exit Function

    End Function
'



' Private Errormethods
    ' Print Error to Immediate
    Private Sub Printt(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty)
        
        Select Case LogMode
            Case LogModeEnum.Standard: Debug.Print GetMessage(Catalog, Index, ErrorValues)
            Case LogModeEnum.Console: HandleOtherLogmode GetMessage(Catalog, Index, ErrorValues)
            Case LogModeEnum.Sheet: HandleOtherLogmode GetMessage(Catalog, Index, ErrorValues)
        End Select
    
    End Sub
    
    ' Print Error as MessageBox
    Private Sub Showw(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty)
        
        Dim Temp As Variant
        Select Case LogMode
            Case LogModeEnum.Standard: Temp = MsgBox(GetMessage(Catalog, Index, ErrorValues), vbExclamation, "ERROR")
            Case LogModeEnum.Console: ' Nothing, as it happens in Printt for the console
            Case LogModeEnum.Sheet: Temp = MsgBox(GetMessage(Catalog, Index, ErrorValues), vbExclamation, "ERROR")
        End Select
        
    End Sub
    
    ' Asks Yes/No Question (No will raise an Error)
    Private Function Ask(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty) As Boolean
        Dim Temp As Variant
        Dim ArrayAnswer() As Variant
        Dim ArrayMessage() As Variant
        ArrayAnswer = Array("y","n")
        ArrayMessage = Array("Answer is Yes", "Answer is No")
        Select Case LogMode
            Case LogModeEnum.Standard: Temp = MsgBox(GetMessage(Catalog, Index, ErrorValues), vbYesNo, "QUESTION")
            Case LogModeEnum.Console: 'Temp = CheckInput(GetMessage(Catalog, Index, ErrorValues), ArrayAnswer, ArrayMessage)
            Case LogModeEnum.Sheet: Temp = MsgBox(GetMessage(Catalog, Index, ErrorValues), vbYesNo, "QUESTION")
        End Select
        If Temp = False Then Ask = IS_ERROR
    End Function

    Private Sub HandleOtherLogmode(Text As String)
        
        Dim i As Long
        If Not LogDestination Is Nothing Then
            Select Case TypeName(LogDestination)
                Case "Range"
                    Do Until LogDestination.Offset(i, 0).Formula = ""
                        i = i + 1
                    Loop
                    LogDestination.Offset(i, 0).Formula = Text
                Case "File"
                    Dim TextStream As Object
                    Set TextStream = LogDestination.OpenAsTextStream(2, 2) '(ForWriting, TristateUseDefault)
                    TextStream.WriteLine Text
                    TextStream.Close
                Case "Console"
                    LogDestination.PrintEnter Text, (&H100000FF - &H10000000)
                Case Else
            End Select
        End If

    End Sub
    
    ' Gets Errormessage
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
    
    ' Runs once to Initialize all Errormessages
    Private Sub ProtInit()
    
        If Initialized = False Then
        ' System-Errors
            ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "std_Error"
            ErrorCatalog(0, 0001) = 1000: ErrorCatalog(1, 0001) = "ErrorCategory doesnt exist"
            ErrorCatalog(0, 0002) = 1000: ErrorCatalog(1, 0002) = "Value isnt available"
            ErrorCatalog(0, 0003) = 1000: ErrorCatalog(1, 0003) = "Value is Empty"
            ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "Value is Nothing"
            ErrorCatalog(0, 0005) = 1000: ErrorCatalog(1, 0005) = "Value Overflow"
            ErrorCatalog(0, 0006) = 1000: ErrorCatalog(1, 0006) = "Value Underflow"
            ErrorCatalog(0, 0007) = 1000: ErrorCatalog(1, 0007) = "Value1 doesnt equal Value2"
            ErrorCatalog(0, 0008) = 1000: ErrorCatalog(1, 0008) = "Value1 is smaller than or equal to Value2"
            ErrorCatalog(0, 0009) = 1000: ErrorCatalog(1, 0009) = "Value1 is bigger than or equal to Value2"
            ErrorCatalog(0, 0010) = 1000: ErrorCatalog(1, 0010) = "Value1 is smaller than Value2"
            ErrorCatalog(0, 0011) = 1000: ErrorCatalog(1, 0011) = "Value1 is bigger than Value2"
            ErrorCatalog(0, 0012) = 1000: ErrorCatalog(1, 0012) = "Value1 is Value2"
            ErrorCatalog(0, 0013) = 1000: ErrorCatalog(1, 0013) = "Several Values are Empty"
            ErrorCatalog(0, 0014) = 1000: ErrorCatalog(1, 0014) = "To many Values arent Empty"
            ErrorCatalog(0, 0015) = 1000: ErrorCatalog(1, 0015) = "Value Is Something"
            ErrorCatalog(0, 0016) = 1000: ErrorCatalog(1, 0016) = "Value1 equals Value2"
            ErrorCatalog(0, 0017) = 1000: ErrorCatalog(1, 0017) = "Value1 is not like Value2"
            ErrorCatalog(0, 0018) = 1000: ErrorCatalog(1, 0018) = "Value1 is like Value2"
            ErrorCatalog(0, 0019) = 1000: ErrorCatalog(1, 0019) = "Value is not a Number"
            ErrorCatalog(0, 0020) = 1000: ErrorCatalog(1, 0020) = "File does not exist"
            ErrorCatalog(0, 0021) = 1000: ErrorCatalog(1, 0021) = "File exists"
            ErrorCatalog(0, 0022) = 1000: ErrorCatalog(1, 0022) = "No active network Connection"
            ErrorCatalog(0, 0023) = 1000: ErrorCatalog(1, 0023) = "No valid IP address found"
            ErrorCatalog(0, 0024) = 1000: ErrorCatalog(1, 0024) = "Unknown error"
            ErrorCatalog(0, 0025) = 1000: ErrorCatalog(1, 0025) = "Active network Connection"
            ErrorCatalog(0, 0026) = 1000: ErrorCatalog(1, 0026) = "valid IP address found"
            ErrorCatalog(0, 0027) = 1000: ErrorCatalog(1, 0027) = "Unable to open Connection"
            ErrorCatalog(0, 0028) = 1000: ErrorCatalog(1, 0028) = "Should not be this Datatype"
            ErrorCatalog(0, 0029) = 1000: ErrorCatalog(1, 0029) = "Should be this Datatype"
            ErrorCatalog(0, 0030) = 1000: ErrorCatalog(1, 0030) = "Unknown Datatype"
            ErrorCatalog(0, 0031) = 1000: ErrorCatalog(1, 0031) = "Value is not Valid"
            ErrorCatalog(0, 0032) = 1000: ErrorCatalog(1, 0032) = "Error-Catalog not a 2D Array with size (1,X)"
            ErrorCatalog(0, 0033) = 1000: ErrorCatalog(1, 0033) = "Error-Catalog not properly initialize"
            ErrorCatalog(0, 0034) = 1000: ErrorCatalog(1, 0034) = "Error-Catalog not an Array"
            ErrorCatalog(0, 0035) = 1000: ErrorCatalog(1, 0035) = "No Internet Connection"
            ErrorCatalog(0, 0036) = 1000: ErrorCatalog(1, 0036) = "Internet Connection is Modem"
            ErrorCatalog(0, 0037) = 1000: ErrorCatalog(1, 0037) = "Internet Connection is Lan"
            ErrorCatalog(0, 0038) = 1000: ErrorCatalog(1, 0038) = "Internet Connection is Proxy"
            ErrorCatalog(0, 0039) = 1000: ErrorCatalog(1, 0039) = "Internet Connection is Offline"
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
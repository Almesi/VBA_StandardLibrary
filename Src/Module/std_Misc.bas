Attribute VB_Name = "std_Workbook"


Option Explicit


Private ErrorCatalog(1, 99) As Variant
Private Initialized As Boolean
Private Handler As New std_Error


Public Function AssignHandler(Optional n_ShowError As Boolean = True, Optional n_LogError As Boolean = True, Optional n_LoggingDestination As Variant, Optional n_ShowDestination As Variant) As Boolean
    On Error GoTo Error
    Let Handler.ShowError = n_ShowError
    Let Handler.LogError = n_LogError
    Set Handler.ShowDestination = n_ShowDestination
    Set Handler.LoggingDestination = n_LoggingDestination
    AssignHandler = True
    Exit Function
    Error:
End Function

Public Property Get IS_ERROR() As Boolean
    IS_ERROR = Handler.IS_ERROR
End Property

Public Function Variable(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR, Optional ThrowError As Boolean = True) As Boolean
    Dim ErrorCode As Long
    If Operator <> EMPTY_ERROR Then
        Select Case UCase(Operator)
            Case "=", "IS"         : If FirstValue =  SecondValue Then ErrorCode = 16: GoTo Error
            Case "<>", "NOT", "!=" : If FirstValue <> SecondValue Then ErrorCode = 07: GoTo Error
            Case "<"               : If FirstValue <  SecondValue Then ErrorCode = 10: GoTo Error
            Case ">"               : If FirstValue >  SecondValue Then ErrorCode = 11: GoTo Error
            Case "=<", "<="        : If FirstValue =< SecondValue Then ErrorCode = 08: GoTo Error
            Case ">=", "=>"        : If FirstValue >= SecondValue Then ErrorCode = 09: GoTo Error
            Case Else
        End Select
    Else
        If FirstValue = Empty                                     Then ErrorCode = 3: GoTo Error
    End If
    If MinValue <> Empty Then         If FirstValue < MinValue    Then ErrorCode = 10: GoTo Error
    If MaxValue <> Empty Then         If FirstValue > MaxValue    Then ErrorCode = 11: GoTo Error
    Exit Function
    Error:
    If ThrowError Then
        Variable = Handler.Handle(ErrorCatalog, ErrorCode, FirstValue, Operator, SecondValue, MinValue, MaxValue)
    Else
        Variable = Handler.IS_ERROR
    End If
End Function

Public Function Object(FirstValue As Object, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Object = EMPTY_ERROR, Optional ThrowError As Boolean = True) As Boolean
    Dim ErrorCode As Long
    If Operator <> EMPTY_ERROR Then
        Select Case UCase(Operator)
            Case "=", "IS"        : If FirstValue     IS SecondValue Then ErrorCode = 16: GoTo Error
            Case "<>", "NOT", "!=": If Not FirstValue IS SecondValue Then ErrorCode = 07: GoTo Error
        End Select
    End If
    If FirstValue Is Nothing                                         Then ErrorCode = 03: GoTo Error
    Exit Function
    Error:
    If ThrowError Then
        Object = Handler.Handle(ErrorCatalog, 04, "FirstValue", Operator, "SecondValue")
    Else
        Object = Handler.IS_ERROR
    End If
End Function

Public Function Strings(Text As String, Optional Operator As String = EMPTY_ERROR, Optional SecondText As String = EMPTY_ERROR, Optional ThrowError As Boolean = True) As Boolean
    Dim ErrorCode As Long
    If Operator <> EMPTY_ERROR Then
        Select Case UCase(Operator)
            Case "=", "IS"           : If Text =        SecondText Then ErrorCode = 16: GoTo Error
            Case "<>", "NOT", "!="   : If Text <>       SecondText Then ErrorCode = 07: GoTo Error
            Case "<"                 : If Text <        SecondText Then ErrorCode = 11: GoTo Error
            Case ">"                 : If Text >        SecondText Then ErrorCode = 10: GoTo Error
            Case "=<", "<="          : If Text =<       SecondText Then ErrorCode = 09: GoTo Error
            Case ">=", "=>"          : If Text >=       SecondText Then ErrorCode = 08: GoTo Error
            Case "LIKE"              : If Text Like     SecondText Then ErrorCode = 18: GoTo Error
            Case "NOT LIKE", "UNLIKE": If Not Text Like SecondText Then ErrorCode = 17: GoTo Error
            Case Else
        End Select
    End If
    If Text = Empty                                                Then ErrorCode = 03: GoTo Error
    Exit Function
    Error:
    If ThrowError Then
        Strings = Handler.Handle(ErrorCatalog, ErrorCode, Text, Operator, SecondText)
    Else
        Strings = Handler.IS_ERROR
    End If
End Function

Public Function Number(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR, Optional ThrowError As Boolean = True) As Boolean
    If IsNumeric(FirstValue) = False Then
        If ThrowError Then
            Number = Handler.Handle(ErrorCatalog, 19, FirstValue, Operator, SecondValue, MinValue, MaxValue)
        Else
            Number = Handler.IS_ERROR
        End If
    Else
        Number = Variable(FirstValue, Operator, SecondValue, MinValue, MaxValue)
    End If
    Exit Function
End Function

Public Function Dates(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR, Optional ThrowError As Boolean = True) As Boolean
    If IsDate(FirstValue) = False Then
        If ThrowError Then
            Dates = Handler.Handle(ErrorCatalog, 19, FirstValue, Operator, SecondValue, MinValue, MaxValue)
        Else
            Dates = Handler.IS_ERROR
        End If
    Else
        Dates = Variable(FirstValue, Operator, SecondValue, MinValue, MaxValue)
    End If
End Function

Public Function File(FilePath As String, Optional ShouldExist As Boolean = True, Optional ThrowError As Boolean = True) As Boolean
    Dim ErrorCode As Long
    If ShouldExist = True Then
        If Dir(FilePath) =  "" Then ErrorCode = 20: GoTo Error
    Else
        If Dir(FilePath) <> "" Then ErrorCode = 21: GoTo Error
    End If
    Exit Function
    Error:
    If ThrowError Then
        File = Handler.Handle(ErrorCatalog, ErrorCode, FilePath)
    Else
        File = std_Error
    End If
End Function

Public Function Connection(Optional Computer As String = ".", Optional ShouldExist As Boolean = True, Optional ThrowError As Boolean = True) As Boolean

    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim ErrorCode As Long
    Set objWMIService = GetObject("winmgmts:\\" & Computer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")

    If ShouldExist  Then
        If colItems.Count = 0 Then
            ErrorCode = 22: GoTo Error
        Else
            For Each objItem In colItems
                If objItem.IPAddress(0) <> "" Then Exit Function
            Next
            ErrorCode = 23: GoTo Error
        End If
    Else
        If colItems.Count <> 0 Then
            ErrorCode = 25: GoTo Error
        Else
            For Each objItem In colItems
                If objItem.IPAddress(0) = "" Then Exit Function
            Next
            ErrorCode = 26: GoTo Error
        End If
    End If
    Exit Function
    If ThrowError Then
        Connection = Handler.Handle(ErrorCatalog, ErrorCatalog, Computer)
    Else
        Connection = Handler.IS_ERROR
    End If

End Function

    ' INTERNET_CONNECTION_MODEM   = &H1
    ' INTERNET_CONNECTION_LAN     = &H2
    ' INTERNET_CONNECTION_PROXY   = &H4
    ' INTERNET_CONNECTION_OFFLINE = &H20

Public Function InternetConnection(Optional ConnectType As Long = 0, Optional ThrowError As Boolean = True) As Boolean
    Dim L As Long
    Dim R As Long
    Dim ErrorCode As Long
    R = InternetGetConnectedState(L, 0&)
    Select Case R
        Case &H00:                         ErrorCode = 35: GoTo Error
        Case &H01: If ConnectType = R Then ErrorCode = 36: GoTo Error
        Case &H02: If ConnectType = R Then ErrorCode = 37: GoTo Error
        Case &H04: If ConnectType = R Then ErrorCode = 38: GoTo Error
        Case &H20: If ConnectType = R Then ErrorCode = 39: GoTo Error
        Case Else
    End Select
    Exit Function
    Error:
    If ThrowError Then
        InternetConnection = Handler.Handle(ErrorCatalog, ErrorCode, ConnectType)
    Else
        InternetConnection = Handler.IS_ERROR
    End If

End Function
                
            

Public Function ConnectToDatabase(DataBasePath As String, Optional ThrowError As Boolean = True) As Boolean
    On Error GoTo Error
    Dim Conn As Object
    Dim ErrorCode As Long
    ErrorCode = 24
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open DataBasePath
    If Conn.State <> 1 Then ErrorCode = 27: GoTo Error
    Exit Function
    Error:
    If ThrowError Then
        ConnectToDatabase = Handler.Handle(ErrorCatalog, ErrorCode, "ConnectToDatabase", DataBasePath)
    Else
        ConnectToDatabase = Handler.IS_ERROR
    End If
End Function

Public Function DataType(Value As Variant, InputType As String, Optional ShouldBe As Boolean = True, Optional ThrowError As Boolean = True) As Boolean
    Dim Inputt As Boolean
    Dim ErrorCode As Long
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
        Case Else: ErrorCode = 30: GoTo Error
    End Select
    If ShouldBe Then
        If Inputt = False Then ErrorCode = 28: GoTo Error
    Else
        If Inputt = True  Then ErrorCode = 29: GoTo Error
    End If 
    Exit Function

    Error:
    If ThrowError Then
        DataType = Handler.Handle(ErrorCatalog, ErrorCode, Value, InputType, ShouldBe)
    Else
        DataType = Handler.IS_ERROR
    End If
End Function


Private Sub ProtInit()
    If Initialized = False Then
        ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "std_Misc"
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
        ErrorCatalog(0, 0032) = 1000: ErrorCatalog(1, 0032) = "PLACEHOLDER"
        ErrorCatalog(0, 0033) = 1000: ErrorCatalog(1, 0033) = "PLACEHOLDER"
        ErrorCatalog(0, 0034) = 1000: ErrorCatalog(1, 0034) = "PLACEHOLDER"
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
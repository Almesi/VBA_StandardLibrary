Attribute VB_Name = "std_Validation"


Option Explicit


Private ErrorCatalog(1, 99) As Variant
Private Initialized As Boolean
Private p_Handler As New std_Error

Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long ) As Long


Public Property Let std_Validation_Handler(n_Handler As std_Error)
    Call ProtInit()
    Set p_Handler = n_Handler
End Property

Public Property Get std_Validation_IS_ERROR() As Boolean
    Call ProtInit()
    std_Validation_IS_ERROR = p_Handler.IS_ERROR
End Property

Public Function std_Validation_Array(Arr As Variant, Optional FirstValue As Variant, Optional Operator As String = Empty, Optional ThrowError As Boolean = True) As Boolean
    Dim i As Long
    Dim Element As Variant
    Select Case True
        Case TypeName(Arr) Like "*()"
            For Each Element In Arr
                std_Validation_Array = std_Validation_Variable(FirstValue, Operator, Element, , , False)
                If std_Validation_Array <> std_Validation_IS_ERROR Then Exit Function
            Next Element
            std_Validation_Array = std_Validation_Variable(FirstValue, Operator, Element, , , True)
        Case TypeName(Arr) = "Range"
            Dim Cell As Range
            For Each Cell In Arr
                std_Validation_Array = std_Validation_Variable(FirstValue, Operator, Cell.Value, , , False)
                If std_Validation_Array <> std_Validation_IS_ERROR Then Exit Function
            Next Cell
            std_Validation_Array = std_Validation_Variable(FirstValue, Operator, Cell.Value, , , True)
        Case TypeName(Arr) = "ComboBox" 
            For Each Element In Arr.List
                std_Validation_Array = std_Validation_Variable(FirstValue, Operator, Element, , , False)
                If std_Validation_Array <> std_Validation_IS_ERROR Then
                    Exit Function
                End If
            Next Element
            std_Validation_Array = std_Validation_Variable(FirstValue, Operator, Element, , , True)
        Case Else
            GoTo Error
    End Select
    Exit Function

    Error:
    std_Validation_Array = p_Handler.Handle(ErrorCatalog, 40, ThrowError, TypeName(Arr))
End Function

Public Function std_Validation_Variable(FirstValue As Variant, Optional Operator As String = Empty, Optional SecondValue As Variant, Optional MinValue As Variant, Optional MaxValue As Variant, Optional ThrowError As Boolean = True) As Boolean
    Dim ErrorCode As Long
    Call ProtInit()
    If IsMissing(FirstValue)                                      Then ErrorCode = 3: GoTo Error
    If Operator <> Empty Then
        Select Case UCase(Operator)
            Case "=", "IS"         : If FirstValue =  SecondValue Then ErrorCode = 16: GoTo Error
            Case "<>", "NOT", "!=" : If FirstValue <> SecondValue Then ErrorCode = 07: GoTo Error
            Case "<"               : If FirstValue <  SecondValue Then ErrorCode = 10: GoTo Error
            Case ">"               : If FirstValue >  SecondValue Then ErrorCode = 11: GoTo Error
            Case "=<", "<="        : If FirstValue =< SecondValue Then ErrorCode = 08: GoTo Error
            Case ">=", "=>"        : If FirstValue >= SecondValue Then ErrorCode = 09: GoTo Error
            Case Else
        End Select
    End If
    If IsMissing(MinValue) = False Then         If FirstValue < MinValue    Then ErrorCode = 10: GoTo Error
    If IsMissing(MaxValue) = False Then         If FirstValue > MaxValue    Then ErrorCode = 11: GoTo Error
    Exit Function
    Error:
    If IsMissing(FirstValue) Then FirstValue = Empty
    If IsMissing(SecondValue) Then SecondValue = Empty
    If IsMissing(MinValue) Then MinValue = Empty
    If IsMissing(MaxValue) Then MaxValue = Empty
    std_Validation_Variable = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, FirstValue, Operator, SecondValue, MinValue, MaxValue)
End Function

Public Function std_Validation_Object(FirstValue As Object, Optional Operator As String = Empty, Optional SecondValue As Object = Empty, Optional ThrowError As Boolean = True) As Boolean
    Dim ErrorCode As Long
    Call ProtInit()
    If Operator <> Empty Then
        Select Case UCase(Operator)
            Case "=", "IS"        : If FirstValue     IS SecondValue Then ErrorCode = 16: GoTo Error
            Case "<>", "NOT", "!=": If Not FirstValue IS SecondValue Then ErrorCode = 07: GoTo Error
        End Select
    End If
    If FirstValue Is Nothing                                         Then ErrorCode = 03: GoTo Error
    Exit Function
    Error:
    std_Validation_Object = p_Handler.Handle(ErrorCatalog, 04, ThrowError, "FirstValue", Operator, "SecondValue")
End Function

Public Function std_Validation_Strings(Text As String, Optional Operator As String = Empty, Optional SecondText As String = Empty, Optional ThrowError As Boolean = True) As Boolean
    Dim ErrorCode As Long
    Call ProtInit()
    If Operator <> Empty Then
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
    std_Validation_Strings = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, Text, Operator, SecondText)
End Function

Public Function std_Validation_Number(FirstValue As Variant, Optional Operator As String = Empty, Optional SecondValue As Variant, Optional MinValue As Variant, Optional MaxValue As Variant, Optional ThrowError As Boolean = True) As Boolean
    Call ProtInit()
    If IsNumeric(FirstValue) = False Then
        If IsMissing(SecondValue) Then SecondValue = Empty
        If IsMissing(MinValue) Then MinValue = Empty
        If IsMissing(MaxValue) Then MaxValue = Empty
        std_Validation_Number = p_Handler.Handle(ErrorCatalog, 19, ThrowError, FirstValue, Operator, SecondValue, MinValue, MaxValue)
    Else
        FirstValue = CNum(FirstValue)
        std_Validation_Number = std_Validation_Variable(FirstValue, Operator, SecondValue, MinValue, MaxValue)
    End If
    Exit Function
End Function

Public Function std_Validation_Dates(FirstValue As Variant, Optional Operator As String = Empty, Optional SecondValue As Variant, Optional MinValue As Variant, Optional MaxValue As Variant, Optional ThrowError As Boolean = True) As Boolean
    Call ProtInit()
    If IsDate(FirstValue) = False Then
        If IsMissing(SecondValue) Then SecondValue = Empty
        If IsMissing(MinValue) Then MinValue = Empty
        If IsMissing(MaxValue) Then MaxValue = Empty
        std_Validation_Dates = p_Handler.Handle(ErrorCatalog, 19, ThrowError, FirstValue, Operator, SecondValue, MinValue, MaxValue)
    Else
        FirstValue = CNum(FirstValue)
        std_Validation_Dates = std_Validation_Variable(FirstValue, Operator, SecondValue, MinValue, MaxValue)
    End If
End Function

Public Function std_Validation_File(FilePath As String, Optional ShouldExist As Boolean = True, Optional ThrowError As Boolean = True) As Boolean
    Dim ErrorCode As Long
    Call ProtInit()
    If ShouldExist = True Then
        If Dir(FilePath) =  "" Then ErrorCode = 20: GoTo Error
    Else
        If Dir(FilePath) <> "" Then ErrorCode = 21: GoTo Error
    End If
    Exit Function
    Error:
    std_Validation_File = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, FilePath)
End Function

Public Function std_Validation_Connection(Optional Computer As String = ".", Optional ShouldExist As Boolean = True, Optional ThrowError As Boolean = True) As Boolean

    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim ErrorCode As Long
    Set objWMIService = GetObject("winmgmts:\\" & Computer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")

    Call ProtInit()
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
    std_Validation_Connection = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, Computer)

End Function

    ' INTERNET_CONNECTION_MODEM   = &H1
    ' INTERNET_CONNECTION_LAN     = &H2
    ' INTERNET_CONNECTION_PROXY   = &H4
    ' INTERNET_CONNECTION_OFFLINE = &H20

Public Function std_Validation_InternetConnection(Optional ConnectType As Long = 0, Optional ThrowError As Boolean = True) As Boolean
    Dim L As Long
    Dim R As Long
    Dim ErrorCode As Long
    Call ProtInit()
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
    std_Validation_InternetConnection = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, ConnectType)

End Function
                
            

Public Function std_Validation_ConnectToDatabase(DataBasePath As String, Optional ThrowError As Boolean = True) As Boolean
    On Error GoTo Error
    Dim Conn As Object
    Dim ErrorCode As Long
    ErrorCode = 24
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open DataBasePath
    If Conn.State <> 1 Then ErrorCode = 27: GoTo Error
    Exit Function
    Error:
    std_Validation_ConnectToDatabase = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, "ConnectToDatabase", DataBasePath)
End Function

Public Function std_Validation_DataType(Value As Variant, InputType As String, Optional ShouldBe As Boolean = True, Optional ThrowError As Boolean = True) As Boolean
    Dim Inputt As Boolean
    Dim ErrorCode As Long
    Call ProtInit()
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
    std_Validation_DataType = p_Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, Value, InputType, ShouldBe)
End Function


Public Function CNum(Value As Variant) As Variant
    On Error Resume Next
    CNum = "x"
    CNum = CDbl(Value)   : If CNum <> "x" Then Exit Function
    CNum = CSng(Value)   : If CNum <> "x" Then Exit Function
    CNum = ClngLng(Value): If CNum <> "x" Then Exit Function
    CNum = CLng(Value)   : If CNum <> "x" Then Exit Function
    CNum = CInt(Value)   : If CNum <> "x" Then Exit Function
    CNum = CByte(Value)  : If CNum <> "x" Then Exit Function
    CNum = Empty
End Function

Private Sub ProtInit()
    If Initialized = False Then
        ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "std_Validation"
        ErrorCatalog(0, 0001) = 1000: ErrorCatalog(1, 0001) = "ErrorCategory doesnt exist"
        ErrorCatalog(0, 0002) = 1000: ErrorCatalog(1, 0002) = "Value isnt available"
        ErrorCatalog(0, 0003) = 1000: ErrorCatalog(1, 0003) = "Value is Empty"
        ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "Value is Nothing"
        ErrorCatalog(0, 0005) = 1000: ErrorCatalog(1, 0005) = "Value Overflow"
        ErrorCatalog(0, 0006) = 1000: ErrorCatalog(1, 0006) = "Value Underflow"
        ErrorCatalog(0, 0007) = 1000: ErrorCatalog(1, 0007) = "One Value doesnt equal another Value"
        ErrorCatalog(0, 0008) = 1000: ErrorCatalog(1, 0008) = "One Value is smaller than or equal to another Value"
        ErrorCatalog(0, 0009) = 1000: ErrorCatalog(1, 0009) = "One Value is bigger than or equal to another Value"
        ErrorCatalog(0, 0010) = 1000: ErrorCatalog(1, 0010) = "One Value is smaller than another Value"
        ErrorCatalog(0, 0011) = 1000: ErrorCatalog(1, 0011) = "One Value is bigger than another Value"
        ErrorCatalog(0, 0012) = 1000: ErrorCatalog(1, 0012) = "One Value is another Value"
        ErrorCatalog(0, 0013) = 1000: ErrorCatalog(1, 0013) = "Several Values are Empty"
        ErrorCatalog(0, 0014) = 1000: ErrorCatalog(1, 0014) = "To many Values arent Empty"
        ErrorCatalog(0, 0015) = 1000: ErrorCatalog(1, 0015) = "Value Is Something"
        ErrorCatalog(0, 0016) = 1000: ErrorCatalog(1, 0016) = "One Value equals another Value"
        ErrorCatalog(0, 0017) = 1000: ErrorCatalog(1, 0017) = "One Value is not like another Value"
        ErrorCatalog(0, 0018) = 1000: ErrorCatalog(1, 0018) = "One Value is like another Value"
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
        ErrorCatalog(0, 0040) = 1000: ErrorCatalog(1, 0040) = "Set not allowed"
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
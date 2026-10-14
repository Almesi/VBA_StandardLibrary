VERSION 1.0 CLASS

BEGIN
    MultiUse = -1
END

Attribute VB_Name = "GenericConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IValueConverter

Private mDataType As CoreValueType

Public Function Create(ByVal DataType As CoreValueType) As GenericConverter
    Dim Result As New GenericConverter
    Call Result.Initialize(DataType)
    Set Create = Result
End Function

Friend Sub Initialize(ByVal DataType As CoreValueType)
    mDataType = DataType
End Sub

Private Function IValueConverter_Convert(ByVal Value As Variant, ByRef Result As Variant) As Boolean
    On Error GoTo ConversionError
    Select Case mDataType
        Case CoreVariant    :  Result = Value
        Case CoreAny        :  Result = Value
        Case CoreString     :  Result = CStr(Value)
        Case CoreChar       :  Result = ConvertChar(Value)
        Case CoreByte       :  Result = CByte(Value)
        Case CoreShort      :  Result = CInt(Value)
        Case CoreInteger    :  Result = CInt(Value)
        Case CoreLong       :  Result = CLng(Value)
        Case CoreLongLong   :  Result = CDbl(Value)
        Case CoreUInteger   :  Result = CInt(Value)
        Case CoreULong      :  Result = CDbl(Value)
        Case CoreULongLong  :  Result = CDbl(Value)
        Case CoreSingle     :  Result = CSng(Value)
        Case CoreDouble     :  Result = CDbl(Value)
        Case CoreExtended   :  Result = CDbl(Value)
        Case CoreDecimal    :  Result = CDec(Value)
        Case CoreDate       :  Result = DateValue(CDate(Value))
        Case CoreTime       :  Result = TimeValue(CDate(Value))
        Case CoreDateTime   :  Result = CDate(Value)
        Case CoreTimeSpan   :  Result = CDbl(Value)
        Case CoreDuration   :  Result = CDbl(Value)
        Case CoreBoolean    :  Result = ConvertBoolean(Value)
        Case CoreNumber     :  Result = CDbl(Value)
        Case CorePercentage :  Result = CDbl(Value)
        Case CoreCurrency   :  Result = CCur(Value)
        Case CoreMoney      :  Result = CCur(Value)
        Case CoreRatio      :  Result = CDbl(Value)
        Case CoreArray      :  If Not IsArray(Value)  Then Exit Function Else Let Result = Value
        Case CoreList       :  If Not IsObject(Value) Then Exit Function Else Set Result = Value
        Case CoreDictionary :  If Not IsObject(Value) Then Exit Function Else Set Result = Value
        Case CoreSet        :  If Not IsObject(Value) Then Exit Function Else Set Result = Value
        Case CoreTuple      :  If Not IsArray(Value)  Then Exit Function Else Let Result = Value
        Case CoreReference  :  If Not IsObject(Value) Then Exit Function Else Set Result = Value
        Case CoreObject     :  If Not IsObject(Value) Then Exit Function Else Set Result = Value
        Case CoreNull       :  Result = Null
        Case CoreEnum       :  Result = CLng(Value)
        Case CoreGuid       :  Result = CStr(Value)
        Case CoreBinary     :  If Not IsArray(Value) Then Exit Function Else Result = Value
        Case CoreChoice     :  Exit Function
        Case Else           :  Exit Function
    End Select
    
    IValueConverter_Convert = True
    Exit Function

    ConversionError:
    IValueConverter_Convert = False
End Function

Private Function ConvertChar(ByVal Value As Variant) As String
    Dim S As String
    S = CStr(Value)
    If Len(S) <> 1 Then
        Err.Raise 13, "GenericConverter", "Value is not a single character."
    End If
    ConvertChar = S
End Function

Private Function ConvertBoolean(ByVal Value As Variant) As Boolean
    If IsBooleanValue(Value) Then
        If VarType(Value) = vbString Then
            ConvertBoolean = (StrComp(CStr(Value), "True", vbTextCompare) = 0)
        Else
            ConvertBoolean = (CDbl(Value) <> 0#)
        End If
        Exit Function
    End If
    Err.Raise 13, "GenericConverter", "Value cannot be converted to Boolean."
End Function

Private Function IsBooleanValue(ByVal Value As Variant) As Boolean
    If IsNull(Value)   Then Exit Function
    If IsObject(Value) Then Exit Function

    Select Case VarType(Value)
        Case vbBoolean                  :  IsBooleanValue = True
        Case vbByte                     :  vbInteger, vbLong, _
             vbSingle, vbDouble, vbCurrency, vbDecimal
             
            IsBooleanValue = _
                (CDbl(Value) = 0# Or CDbl(Value) = 1#)
        Case vbString                   :  IsBooleanValue = _
                StrComp(CStr(Value), "True", vbTextCompare) = 0 Or _
                StrComp(CStr(Value), "False", vbTextCompare) = 0
                
    End Select
    
End Function

Private Property Get IValueConverter_TargetType() As CoreValueType
    IValueConverter_TargetType = mDataType
End Property
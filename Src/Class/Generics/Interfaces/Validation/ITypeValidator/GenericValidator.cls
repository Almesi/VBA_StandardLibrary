VERSION 1.0 CLASS

BEGIN
    MultiUse = -1
END

Attribute VB_Name = "GenericValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements ITypeValidator

Private mDataType As CoreValueType

Public Function Create(ByVal DataType As CoreValueType) As GenericValidator
    Dim Result As New GenericValidator
    
    Call Result.Initialize(DataType)
    
    Set Create = Result
End Function

Friend Sub Initialize(ByVal DataType As CoreValueType)
    mDataType = DataType
End Sub

Private Function ITypeValidator_IsValid(ByVal Value As Variant) As Boolean
    Select Case mDataType
        Case CoreVariant       :  ITypeValidator_IsValid = True
        Case CoreAny           :  ITypeValidator_IsValid = True
        Case CoreString        :  ITypeValidator_IsValid = IsString(Value)
        Case CoreChar          :  ITypeValidator_IsValid = IsChar(Value)
        Case CoreByte          :  ITypeValidator_IsValid = IsByte(Value)
        Case CoreShort         :  ITypeValidator_IsValid = IsShort(Value)
        Case CoreInteger       :  ITypeValidator_IsValid = IsInteger(Value)
        Case CoreLong          :  ITypeValidator_IsValid = IsLong(Value)
        Case CoreLongLong      :  ITypeValidator_IsValid = IsLongLong(Value)
        Case CoreUInteger      :  ITypeValidator_IsValid = IsUInteger(Value)
        Case CoreULong         :  ITypeValidator_IsValid = IsULong(Value)
        Case CoreULongLong     :  ITypeValidator_IsValid = IsULongLong(Value)
        Case CoreSingle        :  ITypeValidator_IsValid = IsSingle(Value)
        Case CoreDouble        :  ITypeValidator_IsValid = IsDouble(Value)
        Case CoreExtended      :  ITypeValidator_IsValid = IsDouble(Value)
        Case CoreDecimal       :  ITypeValidator_IsValid = IsDecimal(Value)
        Case CoreDate          :  ITypeValidator_IsValid = IsDateValue(Value)
        Case CoreTime          :  ITypeValidator_IsValid = IsTime(Value)
        Case CoreDateTime      :  ITypeValidator_IsValid = IsDateValue(Value)
        Case CoreTimeSpan      :  ITypeValidator_IsValid = IsDuration(Value)
        Case CoreDuration      :  ITypeValidator_IsValid = IsDuration(Value)
        Case CoreBoolean       :  ITypeValidator_IsValid = IsBoolean(Value)
        Case CoreNumber        :  ITypeValidator_IsValid = IsNumericValue(Value)
        Case CorePercentage    :  ITypeValidator_IsValid = IsNumericValue(Value)
        Case CoreCurrency      :  ITypeValidator_IsValid = IsNumericValue(Value)
        Case CoreMoney         :  ITypeValidator_IsValid = IsNumericValue(Value)
        Case CoreRatio         :  ITypeValidator_IsValid = IsNumericValue(Value)
        Case CoreArray         :  ITypeValidator_IsValid = IsArray(Value)
        Case CoreList          :  ITypeValidator_IsValid = IsObject(Value)
        Case CoreDictionary    :  ITypeValidator_IsValid = IsObject(Value)
        Case CoreSet           :  ITypeValidator_IsValid = IsObject(Value)
        Case CoreTuple         :  ITypeValidator_IsValid = IsArray(Value)
        Case CoreReference     :  ITypeValidator_IsValid = IsObject(Value)
        Case CoreObject        :  ITypeValidator_IsValid = IsObject(Value)
        Case CoreNull          :  ITypeValidator_IsValid = IsNull(Value)
        Case CoreEnum          :  ITypeValidator_IsValid = IsNumericValue(Value)
        Case CoreGuid          :  ITypeValidator_IsValid = IsGuid(Value)
        Case CoreBinary        :  ITypeValidator_IsValid = IsBinary(Value)
        Case CoreChoice        :  ITypeValidator_IsValid = False
        Case Else              :  ITypeValidator_IsValid = False    
    End Select
End Function

' ============================================================================
' Text
' ============================================================================

Private Function IsString(ByVal Value As Variant) As Boolean
    IsString = (VarType(Value) = vbString)
End Function

Private Function IsChar(ByVal Value As Variant) As Boolean
    If VarType(Value) <> vbString Then Exit Function
    
    IsChar = (Len(Value) = 1)
End Function

' ============================================================================
' Integers
' ============================================================================

Private Function IsByte(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) < 0# Then Exit Function
    If CDbl(Value) > 255# Then Exit Function
    
    IsByte = True
End Function

Private Function IsShort(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) < -32768# Then Exit Function
    If CDbl(Value) > 32767# Then Exit Function
    
    IsShort = True
End Function

Private Function IsInteger(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) < -32768# Then Exit Function
    If CDbl(Value) > 32767# Then Exit Function
    
    IsInteger = True
End Function

Private Function IsLong(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) < -2147483648# Then Exit Function
    If CDbl(Value) > 2147483647# Then Exit Function
    
    IsLong = True
End Function

Private Function IsLongLong(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    
    If CDbl(Value) < -9007199254740992# Then Exit Function
    If CDbl(Value) > 9007199254740992# Then Exit Function
    
    IsLongLong = True
End Function

Private Function IsUInteger(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) < 0# Then Exit Function
    If CDbl(Value) > 65535# Then Exit Function
    
    IsUInteger = True
End Function

Private Function IsULong(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) < 0# Then Exit Function
    If CDbl(Value) > 4294967295# Then Exit Function
    
    IsULong = True
End Function

Private Function IsULongLong(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) < 0# Then Exit Function
    If CDbl(Value) > 9007199254740992# Then Exit Function
    
    IsULongLong = True
End Function

' ============================================================================
' Floating point
' ============================================================================

Private Function IsSingle(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    On Error GoTo InvalidValue
    
    Dim Test As Single
    Test = CSng(Value)
    
    IsSingle = True
    Exit Function
    
    InvalidValue:
    IsSingle = False
End Function

Private Function IsDouble(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    On Error GoTo InvalidValue
    
    Dim Test As Double
    Test = CDbl(Value)
    
    IsDouble = True
    Exit Function
    
    InvalidValue:
    IsDouble = False
End Function

Private Function IsDecimal(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    
    On Error GoTo InvalidValue
    
    Dim Test As Variant
    Test = CDec(Value)
    
    IsDecimal = True
    Exit Function
    
    InvalidValue:
    IsDecimal = False
End Function

' ============================================================================
' Date / Time
' ============================================================================

Private Function IsDateValue(ByVal Value As Variant) As Boolean
    If IsObject(Value) Then Exit Function
    If IsNull(Value) Then Exit Function
    
    IsDateValue = IsDate(Value)
End Function

Private Function IsTime(ByVal Value As Variant) As Boolean
    If Not IsDateValue(Value) Then Exit Function
    
    IsTime = (CDbl(CDate(Value)) >= 0# And _
              CDbl(CDate(Value)) < 1#)
End Function

Private Function IsDuration(ByVal Value As Variant) As Boolean
    IsDuration = IsNumericValue(Value)
End Function

' ============================================================================
' Boolean
' ============================================================================

Private Function IsBoolean(ByVal Value As Variant) As Boolean
    
    If IsNull(Value) Then Exit Function
    If IsObject(Value) Then Exit Function
    
    Select Case VarType(Value)
        Case vbBoolean         :  IsBoolean = True
        Case vbByte            :  vbInteger, vbLong, _
             vbSingle, vbDouble, vbCurrency, vbDecimal
             
            IsBoolean = (CDbl(Value) = 0# Or CDbl(Value) = 1#)
        Case vbString          :  IsBoolean = _
                StrComp(CStr(Value), "True", vbTextCompare) = 0 Or _
                StrComp(CStr(Value), "False", vbTextCompare) = 0
                
    End Select
    
End Function

' ============================================================================
' Numeric abstractions
' ============================================================================

Private Function IsNumericValue(ByVal Value As Variant) As Boolean
    If IsNull(Value) Then Exit Function
    If IsObject(Value) Then Exit Function
    
    IsNumericValue = IsNumeric(Value)
End Function

' ============================================================================
' Special
' ============================================================================

Private Function IsGuid(ByVal Value As Variant) As Boolean
    If VarType(Value) <> vbString Then Exit Function
    
    IsGuid = LooksLikeGuid(CStr(Value))
End Function

Private Function LooksLikeGuid(ByVal Value As String) As Boolean
    Dim S As String
    
    S = Replace(Value, "-", vbNullString)
    
    If Len(S) <> 32 Then Exit Function
    
    LooksLikeGuid = IsHexString(S)
End Function

Private Function IsHexString(ByVal Value As String) As Boolean
    
    Dim I As Long
    Dim C As String
    
    If Len(Value) = 0 Then Exit Function
    
    For I = 1 To Len(Value)
        C = Mid(Value, I, 1)
        
        If InStr(1, "0123456789ABCDEFabcdef", C, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next I
    
    IsHexString = True
End Function

Private Function IsBinary(ByVal Value As Variant) As Boolean
    IsBinary = _
        ((VarType(Value) And vbArray) <> 0) And _
        ((VarType(Value) And vbByte) <> 0)
End Function
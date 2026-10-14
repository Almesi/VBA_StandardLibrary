VERSION 1.0 CLASS

BEGIN
    MultiUse = -1
END

Attribute VB_Name = "GenericChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IValueChecker

Private mDataType As CoreValueType

Public Function Create(ByVal DataType As CoreValueType) As GenericChecker
    Dim Result As New GenericChecker
    
    Call Result.Initialize(DataType)
    
    Set Create = Result
End Function

Friend Sub Initialize(ByVal DataType As CoreValueType)
    mDataType = DataType
End Sub

Private Function IValueChecker_IsValid(ByVal Value As Variant) As Boolean
    Select Case mDataType
        Case CoreVariant       :  IValueChecker_IsValid = True
        Case CoreAny           :  IValueChecker_IsValid = True
        Case CoreString        :  IValueChecker_IsValid = IsString(Value)
        Case CoreChar          :  IValueChecker_IsValid = IsChar(Value)
        Case CoreByte          :  IValueChecker_IsValid = IsByte(Value)
        Case CoreShort         :  IValueChecker_IsValid = IsShort(Value)
        Case CoreInteger       :  IValueChecker_IsValid = IsInteger(Value)
        Case CoreLong          :  IValueChecker_IsValid = IsLong(Value)
        Case CoreLongLong      :  IValueChecker_IsValid = IsLongLong(Value)
        Case CoreUInteger      :  IValueChecker_IsValid = IsUInteger(Value)
        Case CoreULong         :  IValueChecker_IsValid = IsULong(Value)
        Case CoreULongLong     :  IValueChecker_IsValid = IsULongLong(Value)
        Case CoreSingle        :  IValueChecker_IsValid = IsSingle(Value)
        Case CoreDouble        :  IValueChecker_IsValid = IsDouble(Value)
        Case CoreExtended      :  IValueChecker_IsValid = IsDouble(Value)
        Case CoreDecimal       :  IValueChecker_IsValid = IsDecimal(Value)
        Case CoreDate          :  IValueChecker_IsValid = IsDateValue(Value)
        Case CoreTime          :  IValueChecker_IsValid = IsTime(Value)
        Case CoreDateTime      :  IValueChecker_IsValid = IsDateValue(Value)
        Case CoreTimeSpan      :  IValueChecker_IsValid = IsDuration(Value)
        Case CoreDuration      :  IValueChecker_IsValid = IsDuration(Value)
        Case CoreBoolean       :  IValueChecker_IsValid = IsBoolean(Value)
        Case CoreNumber        :  IValueChecker_IsValid = IsNumericValue(Value)
        Case CorePercentage    :  IValueChecker_IsValid = IsNumericValue(Value)
        Case CoreCurrency      :  IValueChecker_IsValid = IsNumericValue(Value)
        Case CoreMoney         :  IValueChecker_IsValid = IsNumericValue(Value)
        Case CoreRatio         :  IValueChecker_IsValid = IsNumericValue(Value)
        Case CoreArray         :  IValueChecker_IsValid = IsArray(Value)
        Case CoreList          :  IValueChecker_IsValid = IsObject(Value)
        Case CoreDictionary    :  IValueChecker_IsValid = IsObject(Value)
        Case CoreSet           :  IValueChecker_IsValid = IsObject(Value)
        Case CoreTuple         :  IValueChecker_IsValid = IsArray(Value)
        Case CoreReference     :  IValueChecker_IsValid = IsObject(Value)
        Case CoreObject        :  IValueChecker_IsValid = IsObject(Value)
        Case CoreNull          :  IValueChecker_IsValid = IsNull(Value)
        Case CoreEnum          :  IValueChecker_IsValid = IsNumericValue(Value)
        Case CoreGuid          :  IValueChecker_IsValid = IsGuid(Value)
        Case CoreBinary        :  IValueChecker_IsValid = IsBinary(Value)
        Case CoreChoice        :  IValueChecker_IsValid = False
        Case Else              :  IValueChecker_IsValid = False    
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
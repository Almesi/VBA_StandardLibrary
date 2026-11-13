Attribute VB_Name = "TypeValidation"
Option Explicit

Public Function IsString(ByVal Value As Variant) As Boolean
    IsString = (VarType(Value) = vbString)
End Function

Public Function IsChar(ByVal Value As Variant) As Boolean
    If VarType(Value) <> vbString Then Exit Function
    IsChar = (Len(Value) = 1)
End Function

Public Function IsByte(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) <  0# Then Exit Function
    If CDbl(Value) >  255# Then Exit Function
    IsByte = True
End Function

Public Function IsShort(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) <  -32768# Then Exit Function
    If CDbl(Value) >  32767# Then Exit Function
    IsShort = True
End Function

Public Function IsInteger(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) <  -32768# Then Exit Function
    If CDbl(Value) >  32767# Then Exit Function
    IsInteger = True
End Function

Public Function IsLong(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) <  -2147483648# Then Exit Function
    If CDbl(Value) >  2147483647# Then Exit Function
    IsLong = True
End Function

Public Function IsLongLong(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) <  -9007199254740992# Then Exit Function
    If CDbl(Value) >  9007199254740992# Then Exit Function
    IsLongLong = True
End Function

Public Function IsUInteger(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) <  0# Then Exit Function
    If CDbl(Value) >  65535# Then Exit Function
    IsUInteger = True
End Function

Public Function IsULong(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) <  0# Then Exit Function
    If CDbl(Value) >  4294967295# Then Exit Function
    IsULong = True
End Function

Public Function IsULongLong(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value)) Then Exit Function
    If CDbl(Value) <  0# Then Exit Function
    If CDbl(Value) >  9007199254740992# Then Exit Function
    IsULongLong = True
End Function

Public Function IsSingle(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    On Error GoTo InvalidValue
    Dim TestValue As Single
    TestValue = CSng(Value)
    IsSingle = True
    Exit Function
    InvalidValue:
    IsSingle = False
End Function

Public Function IsDouble(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    On Error GoTo InvalidValue
    Dim TestValue As Double
    TestValue = CDbl(Value)
    IsDouble = True
    Exit Function
    InvalidValue:
    IsDouble = False
End Function

Public Function IsDecimal(ByVal Value As Variant) As Boolean
    If Not IsNumericValue(Value) Then Exit Function
    On Error GoTo InvalidValue
    Dim TestValue As Variant
    TestValue = CDec(Value)
    IsDecimal = True
    Exit Function
    InvalidValue:
    IsDecimal = False
End Function

Public Function IsDateValue(ByVal Value As Variant) As Boolean
    If IsObject(Value) Then Exit Function
    If IsNull(Value) Then Exit Function
    IsDateValue = IsDate(Value)
End Function

Public Function IsTime(ByVal Value As Variant) As Boolean
    If Not IsDateValue(Value) Then Exit Function
    IsTime = (CDbl(CDate(Value)) >= 0# And CDbl(CDate(Value)) <  1#)
End Function

Public Function IsDuration(ByVal Value As Variant) As Boolean
    IsDuration = IsNumericValue(Value)
End Function

Public Function IsBoolean(ByVal Value As Variant) As Boolean
    If IsNull(Value) Then Exit Function
    If IsObject(Value) Then Exit Function
    Select Case VarType(Value)
        Case vbBoolean
            IsBoolean = True
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsBoolean = (CDbl(Value) = 0# Or CDbl(Value) = 1#)
        Case vbString
            IsBoolean = (StrComp(CStr(Value), "True", vbTextCompare) = 0 Or _
                         StrComp(CStr(Value), "False", vbTextCompare) = 0)
    End Select
End Function

Public Function IsNumericValue(ByVal Value As Variant) As Boolean
    If IsNull(Value) Then Exit Function
    If IsObject(Value) Then Exit Function
    IsNumericValue = IsNumeric(Value)
End Function

Public Function IsGuid(ByVal Value As Variant) As Boolean
    If VarType(Value) <> vbString Then Exit Function
    IsGuid = LooksLikeGuid(CStr(Value))
End Function

Private Function LooksLikeGuid(ByVal Value As String) As Boolean
    Dim TextValue As String
    TextValue = Replace(Value, "-", vbNullString)
    If Len(TextValue) <> 32 Then Exit Function
    LooksLikeGuid = IsHexString(TextValue)
End Function

Private Function IsHexString(ByVal Value As String) As Boolean
    Dim Index As Long
    Dim CharacterValue As String
    If Len(Value) = 0 Then Exit Function
    For Index = 1 To Len(Value)
        CharacterValue = Mid(Value, Index, 1)
        If InStr(1, "0123456789ABCDEFabcdef", CharacterValue, vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next Index
    IsHexString = True
End Function

Public Function IsBinary(ByVal Value As Variant) As Boolean
    IsBinary = ((VarType(Value) And vbArray) <> 0) And ((VarType(Value) And vbByte) <> 0)
End Function

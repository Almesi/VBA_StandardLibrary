Attribute VB_Name = "TypeConversion"
Option Explicit

Public Function ConvertChar(ByVal Value As Variant) As String
    Dim TextValue As String
    TextValue = CStr(Value)
    If Len(TextValue) <> 1 Then
        Err.Raise 13, "CharConverter", "Value is not a single character."
    End If
    ConvertChar = TextValue
End Function

Public Function ConvertBoolean(ByVal Value As Variant) As Boolean
    If IsBoolean(Value) Then
        If VarType(Value) = vbString Then
            ConvertBoolean = (StrComp(CStr(Value), "True", vbTextCompare) = 0)
        Else
            ConvertBoolean = (CDbl(Value) <> 0#)
        End If
        Exit Function
    End If
    Err.Raise 13, "BooleanConverter", "Value cannot be converted to Boolean."
End Function
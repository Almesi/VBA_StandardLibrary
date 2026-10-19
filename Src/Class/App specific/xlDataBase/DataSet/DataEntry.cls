VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mValue As Variant

'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Get Value() As Variant
    If IsObject(mValue) Then
        Set Value = mValue
    Else
        Let Value = mValue
    End If
End Property

Public Property Let Value(ByVal Value As Variant)
    Let mValue = Value
End Property

Public Property Set Value(ByVal Value As Variant)
    Set mValue = Value
End Property

'============================================================
'====================[Creation]==============================
'============================================================

Public Function Create(ByVal Value As Variant) As DataEntry
    Set Create = New DataEntry
    Create.Value = Value
End Function

Public Function Copy(ByVal Other As DataEntry) As DataEntry
    Set Copy = New DataEntry
    Copy.Value = Other.Value
End Function

'============================================================
'====================[Comparison]============================
'============================================================

Public Function Equals(ByVal Other As DataEntry) As Boolean
    If Other Is Nothing Then Exit Function
    Equals = ValuesAreEqual(mValue, Other.Value)
End Function

Public Function EqualsValue(ByVal OtherValue As Variant) As Boolean
    EqualsValue = ValuesAreEqual(mValue, OtherValue)
End Function

Private Function ValuesAreEqual(ByRef LeftValue As Variant, ByRef RightValue As Variant ) As Boolean
    If IsObject(LeftValue) Then
        If Not IsObject(RightValue) Then Exit Function

        ValuesAreEqual = (LeftValue Is RightValue)
        Exit Function
    End If

    If IsObject(RightValue) Then Exit Function

    If IsNull(LeftValue) Or IsNull(RightValue) Then
        ValuesAreEqual = IsNull(LeftValue) And IsNull(RightValue)
        Exit Function
    End If

    If IsEmpty(LeftValue) Or IsEmpty(RightValue) Then
        ValuesAreEqual = IsEmpty(LeftValue) And IsEmpty(RightValue)
        Exit Function
    End If

    ValuesAreEqual = (LeftValue = RightValue)
End Function

Public Function IsGreaterThan(ByVal OtherValue As Variant) As Boolean
    If Not CanCompare(OtherValue) Then Exit Function
    IsGreaterThan = (mValue > OtherValue)
End Function

Public Function IsLessThan(ByVal OtherValue As Variant) As Boolean
    If Not CanCompare(OtherValue) Then Exit Function
    IsLessThan = (mValue < OtherValue)
End Function

Public Function IsGreaterThanOrEqual(ByVal OtherValue As Variant ) As Boolean
    If Not CanCompare(OtherValue) Then Exit Function
    IsGreaterThanOrEqual = (mValue >= OtherValue)
End Function

Public Function IsLessThanOrEqual(ByVal OtherValue As Variant ) As Boolean
    If Not CanCompare(OtherValue) Then Exit Function
    IsLessThanOrEqual = (mValue <= OtherValue)
End Function

Private Function CanCompare(ByRef OtherValue As Variant) As Boolean
    If IsObject(mValue) Or IsObject(OtherValue) Then Exit Function
    If IsNull(mValue)   Or IsNull(OtherValue)   Then Exit Function
    If IsEmpty(mValue)  Or IsEmpty(OtherValue)  Then Exit Function

    CanCompare = True
End Function
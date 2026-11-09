VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ChoiceValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements ITypeValidator

Private mChoice As Collection

Public Function Create(ByVal Choice As Collection) As ChoiceValidator
    Dim Result As New ChoiceValidator
    Call Result.Initialize(Choice)
    Set Create = Result
End Function

Friend Sub Initialize(ByVal Choice As Collection)
    Set mChoice = Choice
End Sub

Private Function ITypeValidator_IsValid(ByVal Value As Variant) As Boolean
    Dim Item As Variant
    
    For Each Item In mChoice
        If StrComp(CStr(Item), CStr(Value), vbTextCompare) = 0 Then
            ITypeValidator_IsValid = True
            Exit Function
        End If
    Next Item
End Function
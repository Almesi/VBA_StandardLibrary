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

Implements IValidator

Private mChoice As Collection

Public Function Create(ByVal Choice As Collection) As ChoiceValidator
    Dim Result As New ChoiceValidator
    Call Result.Initialize(Choice)
    Set Create = Result
End Function

Friend Sub Initialize(ByVal Choice As Collection)
    Set mChoice = Choice
End Sub

Private Function IValidator_IsValid(ByVal Value As Variant) As Boolean
    Dim Item As Variant
    
    For Each Item In mChoice
        If StrComp(CStr(Item), CStr(Value), vbTextCompare) = 0 Then
            IValidator_IsValid = True
            Exit Function
        End If
    Next Item
End Function

Private Function IValidator_Identity() As String
    Dim Message As String
    Dim Item As Variant
    For Each Item In mChoice
        Message = Message & "|" & Item
    Next Item
    IValidator_Identity = Mid(Message, 2, Len(Message) - 1)
End Function
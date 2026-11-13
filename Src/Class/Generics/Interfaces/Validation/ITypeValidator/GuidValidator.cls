VERSION 1.0 CLASS

BEGIN
    MultiUse = -1
END

Attribute VB_Name = "GuidValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IValidator

Public Function Create() As GuidValidator
    Dim Result As New GuidValidator
    Call Result.Initialize
    Set Create = Result
End Function

Friend Sub Initialize()
End Sub

Private Function IValidator_IsValid(ByVal Value As Variant) As Boolean
    IValidator_IsValid = IsGuid(Value)
End Function

Private Function IValidator_Identity() As String
    IValidator_Identity = "GUID"
End Function
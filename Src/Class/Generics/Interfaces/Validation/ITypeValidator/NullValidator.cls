VERSION 1.0 CLASS

BEGIN
    MultiUse = -1
END

Attribute VB_Name = "NullValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IValidator

Public Function Create() As NullValidator
    Dim Result As New NullValidator
    Call Result.Initialize
    Set Create = Result
End Function

Friend Sub Initialize()
End Sub

Private Function IValidator_IsValid(ByVal Value As Variant) As Boolean
    IValidator_IsValid = IsNull(Value)
End Function

Private Function IValidator_Identity() As String
    IValidator_Identity = "Null"
End Function
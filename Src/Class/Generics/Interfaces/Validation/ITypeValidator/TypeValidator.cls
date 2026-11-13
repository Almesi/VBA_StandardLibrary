VERSION 1.0 CLASS

BEGIN
    MultiUse = -1
END

Attribute VB_Name = "TypeValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IValidator

Private mClassName As String

Public Function CreateFromVariant(ByVal Variable As Variant) As TypeValidator
    Set CreateFromVariant = Create(TypeName(Variable))
End Function

Public Function Create(ByVal ClassName As String) As TypeValidator
    Dim Result As New TypeValidator
    Call Result.Initialize(ClassName)
    Set Create = Result
End Function

Friend Sub Initialize(ByVal ClassName As String)
    mClassName = ClassName
End Sub

Private Function IValidator_IsValid(ByVal Value As Variant) As Boolean
    IValidator_IsValid = (mClassName = TypeName(Value))
End Function

Private Function IValidator_Identity() As String
    IValidator_Identity = mClassName
End Function
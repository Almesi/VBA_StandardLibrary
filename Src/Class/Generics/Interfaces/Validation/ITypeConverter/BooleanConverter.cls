VERSION 1.0 CLASS

BEGIN
    MultiUse = -1
END

Attribute VB_Name = "BooleanConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements ITypeConverter

Public Function Create() As BooleanConverter
    Dim Result As New BooleanConverter
    Call Result.Initialize
    Set Create = Result
End Function

Friend Sub Initialize()
End Sub

Private Function ITypeConverter_Convert(ByVal Value As Variant, ByRef Result As Variant) As Boolean
    On Error GoTo ConversionError

    Result = TypeConversion.ConvertBoolean(Value)

    ITypeConverter_Convert = True
    Exit Function

    ConversionError:
    ITypeConverter_Convert = False
End Function

Private Property Get ITypeConverter_TargetType() As CoreValueType
    ITypeConverter_TargetType = CoreBoolean
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ChoiceConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IValueConverter

Private mDataType As CoreValueType

Public Function Create(ByVal DataType As CoreValueType) As ChoiceConverter
    Dim Result As New ChoiceConverter
    Call Result.Initialize(DataType)
    Set Create = Result
End Function

Friend Sub Initialize(ByVal DataType As CoreValueType)
    mDataType = DataType
End Sub

Public Function IValueConverter_Convert(ByVal Value As Variant, ByRef Result As Variant) As Boolean
    On Error GoTo ConversionError
    
    Result = CStr(Value)
    
    IValueConverter_Convert = True
    Exit Function

ConversionError:
    IValueConverter_Convert = False
End Function

Private Property Get IValueConverter_TargetType() As CoreValueType
    TargetType = mDataType
End Property
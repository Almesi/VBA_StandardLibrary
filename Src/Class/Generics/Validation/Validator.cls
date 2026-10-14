VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mChecker   As IValueChecker
Private mConverter As IValueConverter

Public Function Create(ByVal Checker As IValueChecker, ByVal Converter As IValueConverter) As Validator
    Dim Result As New Validator
    Call Result.Initialize(Checker, Converter)
    Set Create = Result
End Function

Friend Sub Initialize(ByVal Checker As IValueChecker, ByVal Converter As IValueConverter)
    Set mChecker   = Checker
    Set mConverter = Converter
End Sub

Public Function Validate(ByVal Value As Variant, ByVal Output As ValidatedValue) As ValidationResult
    Dim ValueIsValid As Boolean
    ValueIsValid = IsValid(Value)
    If Not ValueIsValid Then
        Set Validate = ValidationResult.CreateFromBool(ValueIsValid, "Checker Failed: " & DataType)
        Exit Function
    End If

    Dim ConvertedValue   As Variant
    Dim ValueConvertable As Boolean
    ValueConvertable = Convert(Value, ConvertedValue)
    Set Validate = ValidationResult.CreateFromBool(ValueConvertable, "Conversion Failed: " & DataType)
        
    Set Output = ValidatedValue.Create(ConvertedValue, TargetType)
End Function

Public Function IsValid(ByVal Value As Variant) As Boolean
    IsValid = mChecker.IsValid(Value)
End Function

Public Function Convert(ByVal Value As Variant, ByRef Result As Variant) As Boolean
    Convert = mConverter.Convert(Value, Result)
End Function

Public Property Get DataType() As CoreValueType
    DataType = mConverter.TargetType
End Property
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

Private mChecker   As ITypeValidator
Private mConverter As ITypeConverter

Public Function CreateFromCoreValueType(ByVal ValueType As CoreValueType) As Validator
    Select Case ValueType
        Case CoreString     : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreString)      , GenericConverter.Create(CoreString))
        Case CoreChar       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreChar)        , GenericConverter.Create(CoreChar))
        Case CoreByte       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreByte)        , GenericConverter.Create(CoreByte))
        Case CoreShort      : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreShort)       , GenericConverter.Create(CoreShort))
        Case CoreInteger    : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreInteger)     , GenericConverter.Create(CoreInteger))
        Case CoreLong       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreLong)        , GenericConverter.Create(CoreLong))
        Case CoreLongLong   : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreLongLong)    , GenericConverter.Create(CoreLongLong))
        Case CoreUInteger   : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreUInteger)    , GenericConverter.Create(CoreUInteger))
        Case CoreULong      : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreULong)       , GenericConverter.Create(CoreULong))
        Case CoreULongLong  : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreULongLong)   , GenericConverter.Create(CoreULongLong))
        Case CoreSingle     : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreSingle)      , GenericConverter.Create(CoreSingle))
        Case CoreDouble     : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreDouble)      , GenericConverter.Create(CoreDouble))
        Case CoreExtended   : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreExtended)    , GenericConverter.Create(CoreExtended))
        Case CoreDecimal    : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreDecimal)     , GenericConverter.Create(CoreDecimal))
        Case CoreDate       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreDate)        , GenericConverter.Create(CoreDate))
        Case CoreTime       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreTime)        , GenericConverter.Create(CoreTime))
        Case CoreDateTime   : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreDateTime)    , GenericConverter.Create(CoreDateTime))
        Case CoreTimeSpan   : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreTimeSpan)    , GenericConverter.Create(CoreTimeSpan))
        Case CoreDuration   : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreDuration)    , GenericConverter.Create(CoreDuration))
        Case CoreBoolean    : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreBoolean)     , GenericConverter.Create(CoreBoolean))
        Case CoreNumber     : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreNumber)      , GenericConverter.Create(CoreNumber))
        Case CorePercentage : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CorePercentage)  , GenericConverter.Create(CorePercentage))
        Case CoreCurrency   : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreCurrency)    , GenericConverter.Create(CoreCurrency))
        Case CoreMoney      : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreMoney)       , GenericConverter.Create(CoreMoney))
        Case CoreRatio      : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreRatio)       , GenericConverter.Create(CoreRatio))
        Case CoreArray      : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreArray)       , GenericConverter.Create(CoreArray))
        Case CoreList       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreList)        , GenericConverter.Create(CoreList))
        Case CoreDictionary : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreDictionary)  , GenericConverter.Create(CoreDictionary))
        Case CoreSet        : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreSet)         , GenericConverter.Create(CoreSet))
        Case CoreTuple      : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreTuple)       , GenericConverter.Create(CoreTuple))
        Case CoreReference  : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreReference)   , GenericConverter.Create(CoreReference))
        Case CoreObject     : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreObject)      , GenericConverter.Create(CoreObject))
        Case CoreNull       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreNull)        , GenericConverter.Create(CoreNull))
        Case CoreVariant    : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreVariant)     , GenericConverter.Create(CoreVariant))
        Case CoreAny        : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreAny)         , GenericConverter.Create(CoreAny))
        'Case CoreEnum       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreEnum)        , GenericConverter.Create(CoreEnum))
        'Case CoreGuid       : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreGuid)        , GenericConverter.Create(CoreGuid))
        'Case CoreBinary     : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreBinary)      , GenericConverter.Create(CoreBinary))
        'Case CoreChoice     : Set CreateFromCoreValueType = Validator.Create(GenericChecker.Create(CoreChoice)      , GenericConverter.Create(CoreChoice))
    End Select
End Function

Public Function Create(ByVal Checker As ITypeValidator, ByVal Converter As ITypeConverter) As Validator
    Dim Result As New Validator
    Call Result.Initialize(Checker, Converter)
    Set Create = Result
End Function

Friend Sub Initialize(ByVal Checker As ITypeValidator, ByVal Converter As ITypeConverter)
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
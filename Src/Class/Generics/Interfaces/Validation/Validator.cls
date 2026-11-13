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

Public Enum CoreValueType

    ' Text
    CoreString
    CoreChar

    ' Integers
    CoreByte
    CoreShort
    CoreInteger
    CoreLong
    CoreLongLong
    CoreUInteger
    CoreULong
    CoreULongLong

    ' Floating point
    CoreSingle
    CoreDouble

    ' Decimal / precise numbers
    CoreDecimal

    ' Time and Date
    CoreDate
    CoreTime

    ' Boolean
    CoreBoolean

    ' Number abstractions
    CoreNumber
    CorePercentage
    CoreCurrency
    CoreMoney
    CoreRatio

    ' Collections
    CoreArray

    ' References / objects
    CoreObject
    CoreNull

    ' Dynamic / language-level
    CoreVariant
    CoreAny

    ' Special
    CoreGuid
    CoreBinary

    ' Constrained values
    CoreChoice

End Enum

Private mChecker   As IValidator
Private mConverter As ITypeConverter

Public Function CreateFromCoreValueType(ByVal ValueType As CoreValueType) As Validator
    Select Case ValueType
        Case CoreString     : Set CreateFromCoreValueType = Validator.Create(StringValidator.Create(CoreString)          , StringConverter.Create(CoreString))
        Case CoreChar       : Set CreateFromCoreValueType = Validator.Create(CharValidator.Create(CoreChar)              , CharConverter.Create(CoreChar))
        Case CoreByte       : Set CreateFromCoreValueType = Validator.Create(ByteValidator.Create(CoreByte)              , ByteConverter.Create(CoreByte))
        Case CoreShort      : Set CreateFromCoreValueType = Validator.Create(ShortValidator.Create(CoreShort)            , ShortConverter.Create(CoreShort))
        Case CoreInteger    : Set CreateFromCoreValueType = Validator.Create(IntegerValidator.Create(CoreInteger)        , IntegerConverter.Create(CoreInteger))
        Case CoreLong       : Set CreateFromCoreValueType = Validator.Create(LongValidator.Create(CoreLong)              , LongConverter.Create(CoreLong))
        Case CoreLongLong   : Set CreateFromCoreValueType = Validator.Create(LongLongValidator.Create(CoreLongLong)      , LongLongConverter.Create(CoreLongLong))
        Case CoreUInteger   : Set CreateFromCoreValueType = Validator.Create(UIntegerValidator.Create(CoreUInteger)      , UIntegerConverter.Create(CoreUInteger))
        Case CoreULong      : Set CreateFromCoreValueType = Validator.Create(ULongValidator.Create(CoreULong)            , ULongConverter.Create(CoreULong))
        Case CoreULongLong  : Set CreateFromCoreValueType = Validator.Create(ULongLongValidator.Create(CoreULongLong)    , ULongLongConverter.Create(CoreULongLong))
        Case CoreSingle     : Set CreateFromCoreValueType = Validator.Create(SingleValidator.Create(CoreSingle)          , SingleConverter.Create(CoreSingle))
        Case CoreDouble     : Set CreateFromCoreValueType = Validator.Create(DoubleValidator.Create(CoreDouble)          , DoubleConverter.Create(CoreDouble))
        Case CoreDecimal    : Set CreateFromCoreValueType = Validator.Create(DecimalValidator.Create(CoreDecimal)        , DecimalConverter.Create(CoreDecimal))
        Case CoreDate       : Set CreateFromCoreValueType = Validator.Create(DateValidator.Create(CoreDate)              , DateConverter.Create(CoreDate))
        Case CoreTime       : Set CreateFromCoreValueType = Validator.Create(TimeValidator.Create(CoreTime)              , TimeConverter.Create(CoreTime))
        Case CoreBoolean    : Set CreateFromCoreValueType = Validator.Create(BooleanValidator.Create(CoreBoolean)        , BooleanConverter.Create(CoreBoolean))
        Case CoreNumber     : Set CreateFromCoreValueType = Validator.Create(NumberValidator.Create(CoreNumber)          , NumberConverter.Create(CoreNumber))
        Case CorePercentage : Set CreateFromCoreValueType = Validator.Create(PercentageValidator.Create(CorePercentage)  , PercentageConverter.Create(CorePercentage))
        Case CoreCurrency   : Set CreateFromCoreValueType = Validator.Create(CurrencyValidator.Create(CoreCurrency)      , CurrencyConverter.Create(CoreCurrency))
        Case CoreMoney      : Set CreateFromCoreValueType = Validator.Create(MoneyValidator.Create(CoreMoney)            , MoneyConverter.Create(CoreMoney))
        Case CoreRatio      : Set CreateFromCoreValueType = Validator.Create(RatioValidator.Create(CoreRatio)            , RatioConverter.Create(CoreRatio))
        Case CoreArray      : Set CreateFromCoreValueType = Validator.Create(ArrayValidator.Create(CoreArray)            , ArrayConverter.Create(CoreArray))
        Case CoreObject     : Set CreateFromCoreValueType = Validator.Create(ObjectValidator.Create(CoreObject)          , ObjectConverter.Create(CoreObject))
        Case CoreNull       : Set CreateFromCoreValueType = Validator.Create(NullValidator.Create(CoreNull)              , NullConverter.Create(CoreNull))
        Case CoreVariant    : Set CreateFromCoreValueType = Validator.Create(VariantValidator.Create(CoreVariant)        , VariantConverter.Create(CoreVariant))
        Case CoreAny        : Set CreateFromCoreValueType = Validator.Create(AnyValidator.Create(CoreAny)                , AnyConverter.Create(CoreAny))
        Case CoreGuid       : Set CreateFromCoreValueType = Validator.Create(GuidValidator.Create(CoreGuid)              , GuidConverter.Create(CoreGuid))
        Case CoreBinary     : Set CreateFromCoreValueType = Validator.Create(BinaryValidator.Create(CoreBinary)          , BinaryConverter.Create(CoreBinary))
        Case CoreChoice     : Set CreateFromCoreValueType = Validator.Create(ChoiceValidator.Create(CoreChoice)          , ChoiceConverter.Create(CoreChoice))
    End Select
End Function

Public Function Create(ByVal Checker As IValidator, ByVal Converter As ITypeConverter) As Validator
    Dim Result As New Validator
    Call Result.Initialize(Checker, Converter)
    Set Create = Result
End Function

Friend Sub Initialize(ByVal Checker As IValidator, ByVal Converter As ITypeConverter)
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
Attribute VB_Name = "ValidatorEnums"

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
    CoreExtended

    ' Decimal / precise numbers
    CoreDecimal

    ' Time and Date
    CoreDate
    CoreTime
    CoreDateTime
    CoreTimeSpan
    CoreDuration

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
    CoreList
    CoreDictionary
    CoreSet
    CoreTuple

    ' References / objects
    CoreReference
    CoreObject
    CoreNull

    ' Dynamic / language-level
    CoreVariant
    CoreAny

    ' Special
    CoreEnum
    CoreGuid
    CoreBinary

    ' Constrained values
    CoreChoice

End Enum
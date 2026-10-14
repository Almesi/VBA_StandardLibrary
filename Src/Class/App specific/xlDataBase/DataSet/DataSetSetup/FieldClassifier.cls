VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "FieldClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mRules As Collection

'============================================================
'====================[Creation]==============================
'============================================================

Private Sub Class_Initialize()
    Set mRules = New Collection
End Sub

Public Function CreateFromDataSet(ByVal SourceDataSet As DataSet, ByVal ChoiceCollection As ChoiceCollector) As FieldClassifier
    
    Dim Classifier As FieldClassifier
    Dim Rule       As FieldClassificationRule
    Dim Row        As RowEntry
    
    Set Classifier = New FieldClassifier
    
    For Each Row In SourceDataSet.Rows
        
        Dim Name     As String
        Dim DataType As CoreValueType
        Dim ChoiceName As String
        
        Name = CStr(Row.ByName("Name").DataEntry.Value)
        DataType = ParseCoreValueType(CStr(Row.ByName("DataType").DataEntry.Value))
        ChoiceName = CStr(Row.ByName("Choice").DataEntry.Value)
        
        Dim Validator As Validator
        Set Validator = MakeValidator(ChoiceName, DataType, ChoiceCollection)
        
        Call Classifier.AddRule( _
            Name, _
            CStr(Row.ByName("Pattern").DataEntry.Value), _
            CLng(Row.ByName("Priority").DataEntry.Value), _
            DataType, _
            Validator)
            
    Next Row
    
    Set CreateFromDataSet = Classifier
End Function

'============================================================
'====================[Getters]===============================
'============================================================

Public Property Get RuleCount() As Long
    RuleCount = mRules.Count
End Property

'============================================================
'====================[Functions]=============================
'============================================================

Public Sub AddRule(ByVal Name As String, ByVal Pattern As String, ByVal Priority As Long, ByVal DataType As CoreValueType, ByVal Validator As Validator)
    Dim Rule As FieldClassificationRule
    Set Rule = New FieldClassificationRule
    With Rule
        Let .Name = Name
        Let .Pattern = Pattern
        Let .Priority = Priority
        Let .DataType = DataType
        Set .Validator = Validator
    End With
    Call mRules.Add(Rule)
End Sub

Public Function Classify(ByVal FieldName As String) As FieldDefinition
    Dim Rule     As FieldClassificationRule
    Dim BestRule As FieldClassificationRule
    
    For Each Rule In mRules
        If Rule.Matches(FieldName) Then
            If BestRule Is Nothing Then
                Set BestRule = Rule
            ElseIf Rule.Priority < BestRule.Priority Then
                Set BestRule = Rule
            End If
        End If
    Next Rule

    Set Classify = New FieldDefinition
    With Classify
        .Name = FieldName
        If Not BestRule Is Nothing Then
            Let .DataType = BestRule.DataType
            Set .Validator = BestRule.Validator
        Else
            Let .DataType = CoreVariant
            Set .Validator = Validator.Create(GenericChecker.Create(vbVariant), GenericConverter.Create(vbVariant))
        End If
    End With
End Function

'============================================================
'====================[Helpers]===============================
'============================================================

Private Function MakeValidator(ByVal Name As String, ByVal DataType As CoreValueType, ByVal ChoiceCollection As ChoiceCollector) As Validator
    Dim Checker   As IValueChecker
    Dim Converter As IValueConverter
    Select Case DataType
        Case CoreChoice
            Dim ChoiceValues As Collection
            Set ChoiceValues = ChoiceCollection.GetChoices(Name)
            Set Checker = ChoiceChecker.Create(ChoiceValues)
            Set Converter = ChoiceConverter.Create(DataType)
        Case Else
            Set Checker = GenericChecker.Create(DataType)
            Set Converter = GenericConverter.Create(DataType)
    End Select
    Set MakeValidator = Validator.Create(Checker, Converter)
End Function

Private Function ParseCoreValueType(ByVal Value As String) As CoreValueType
    Select Case LCase(Trim(Value))
        Case "variant"     : ParseCoreValueType = CoreVariant
        Case "any"         : ParseCoreValueType = CoreAny
        Case "string"      : ParseCoreValueType = CoreString
        Case "char"        : ParseCoreValueType = CoreChar
        Case "byte"        : ParseCoreValueType = CoreByte
        Case "short"       : ParseCoreValueType = CoreShort
        Case "integer"     : ParseCoreValueType = CoreInteger
        Case "long"        : ParseCoreValueType = CoreLong
        Case "longlong"    : ParseCoreValueType = CoreLongLong
        Case "uinteger"    : ParseCoreValueType = CoreUInteger
        Case "ulong"       : ParseCoreValueType = CoreULong
        Case "ulonglong"   : ParseCoreValueType = CoreULongLong
        Case "single"      : ParseCoreValueType = CoreSingle
        Case "double"      : ParseCoreValueType = CoreDouble
        Case "extended"    : ParseCoreValueType = CoreExtended
        Case "decimal"     : ParseCoreValueType = CoreDecimal
        Case "date"        : ParseCoreValueType = CoreDate
        Case "time"        : ParseCoreValueType = CoreTime
        Case "datetime"    : ParseCoreValueType = CoreDateTime
        Case "timespan"    : ParseCoreValueType = CoreTimeSpan
        Case "duration"    : ParseCoreValueType = CoreDuration
        Case "boolean"     : ParseCoreValueType = CoreBoolean
        Case "number"      : ParseCoreValueType = CoreNumber
        Case "percentage"  : ParseCoreValueType = CorePercentage
        Case "currency"    : ParseCoreValueType = CoreCurrency
        Case "money"       : ParseCoreValueType = CoreMoney
        Case "ratio"       : ParseCoreValueType = CoreRatio
        Case "array"       : ParseCoreValueType = CoreArray
        Case "list"        : ParseCoreValueType = CoreList
        Case "dictionary"  : ParseCoreValueType = CoreDictionary
        Case "set"         : ParseCoreValueType = CoreSet
        Case "tuple"       : ParseCoreValueType = CoreTuple
        Case "reference"   : ParseCoreValueType = CoreReference
        Case "object"      : ParseCoreValueType = CoreObject
        Case "null"        : ParseCoreValueType = CoreNull
        Case "enum"        : ParseCoreValueType = CoreEnum
        Case "guid"        : ParseCoreValueType = CoreGuid
        Case "binary"      : ParseCoreValueType = CoreBinary
        Case "choice"      : ParseCoreValueType = CoreChoice
        
        Case Else
            
            Err.Raise _
                vbObjectError + 1000, _
                "FieldClassifier.ParseCoreValueType", _
                "Unknown field data type: '" & Value & "'"
            
    End Select
End Function
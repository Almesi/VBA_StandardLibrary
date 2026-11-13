VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "FieldClassificationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mName       As String
Private mPattern    As String
Private mPriority   As Long
Private mDataType   As CoreValueType
Private mValidator  As Validator

'============================================================
'====================[Getters]===============================
'============================================================


Public Property Let Name(ByVal      Value As String)        : Let mName      = Value      : End Property
Public Property Let Pattern(ByVal   Value As String)        : Let mPattern   = Value      : End Property
Public Property Let Priority(ByVal  Value As Long)          : Let mPriority  = Value      : End Property
Public Property Let DataType(ByVal  Value As CoreValueType) : Let mDataType  = Value      : End Property
Public Property Set Validator(ByVal Value As Validator)     : Set mValidator = Value      : End Property

Public Property Get Name()                As String         : Let Name       = mName      : End Property
Public Property Get Pattern()             As String         : Let Pattern    = mPattern   : End Property
Public Property Get Priority()            As Long           : Let Priority   = mPriority  : End Property
Public Property Get DataType()            As CoreValueType  : Let DataType   = mDataType  : End Property
Public Property Get Validator()           As Validator      : Set Validator  = mValidator : End Property

'============================================================
'====================[Functions]=============================
'============================================================

Public Function Matches(ByVal FieldName As String) As Boolean
    Dim Regex As RegExp
    Set Regex = New RegExp

    With Regex
        .Pattern    = mPattern
        .IgnoreCase = True
        .Global     = False
    End With
    Matches = Regex.Test(FieldName)
End Function
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "FieldDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mName      As String
Private mDataType  As VbVarType
Private mValidator As Validator

'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Let Name(ByVal      Value As String)     : Let mName      = Value      : End Property
Public Property Let DataType(ByVal  Value As VbVarType)  : Let mDataType  = Value      : End Property
Public Property Set Validator(ByVal Value As Validator) : Set mValidator = Value      : End Property

Public Property Get Name()                As String      : Let Name       = mName      : End Property
Public Property Get DataType()            As VbVarType   : Let DataType   = mDataType  : End Property
Public Property Get Validator()           As Validator  : Set Validator  = mValidator : End Property

Public Function Create(ByVal nName As String, ByVal nDataType As VbVarType, ByVal nValidator As Validator) As FieldDefinition
    Set Create = New FieldDefinition
    Create.Name       = nName
    Create.DataType   = nDataType
    Create.Validator  = nValidator
End Function

'============================================================
'====================[Creation]==============================
'============================================================

Public Function Copy(ByVal Other As FieldDefinition) As FieldDefinition
    Set Copy = New FieldDefinition
    Let Copy.Name      = Other.Name
    Let Copy.DataType  = Other.DataType
    Set Copy.Validator = Other.Validator
End Function

'============================================================
'====================[Functions]=============================
'============================================================
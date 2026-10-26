
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Enum QueryValueKind
    QueryValueKindNull     = 0
    QueryValueKindScalar   = 1
    QueryValueKindSequence = 2
    QueryValueKindDataSet  = 3
End Enum

Private mKind  As QueryValueKind
Private mValue As Variant

Public Function Create(ByVal NewValue As Variant) As QueryValue
    Dim KindType As QueryValueKind
    Select Case True
        Case VarType(NewValue) Like "DataSet"    : KindType = QueryValueKindDataSet
        Case VarType(NewValue) Like "Collectoin" : KindType = QueryValueKindSequence
        Case VarType(NewValue) Like "Null"       : KindType = QueryValueKindNull
        Case Else                                : KindType = QueryValueKindScalar
    End Select
    Set Create = CreateNormal(KindType, NewValue)
End Function

Public Function CreateNormal(ByVal QueryKind As QueryValueKind, ByVal NewValue As Variant) As QueryValue
    Set CreateNormal = New QueryValue
    Call CreateNormal.Assign(QueryKind, NewValue)
End Function

Public Property Get Kind() As QueryValueKind
    Kind = mKind
End Property

Public Property Get IsNull() As Boolean
    IsNull = (mKind = QueryValueKindNull)
End Property

Public Property Get IsScalar() As Boolean
    IsScalar = (mKind = QueryValueKindScalar)
End Property

Public Property Get IsSequence() As Boolean
    IsSequence = (mKind = QueryValueKindSequence)
End Property

Public Property Get IsDataSet() As Boolean
    IsDataSet = (mKind = QueryValueKindDataSet)
End Property

Public Property Get Value() As Variant
    If IsObject(mValue) Then
        Set Value = mValue
    Else
        Let Value = mValue
    End If
End Property

Public Sub Assign(ByVal QueryKind As QueryValueKind, ByVal NewValue As Variant)
    Let mKind  = QueryKind
    If IsObject(NewValue) Then
        Set mValue = NewValue
    Else
        Let mValue = NewValue
    End If
End Sub
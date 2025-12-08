VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_IncludeSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private pValue As Variant
Private pType  As String

Public Function Create(ByVal nValue As Variant) As std_IncludeSource
    Set Create = New std_IncludeSource
    Create.Value = nValue
End Function

Public Property Let Value(ByVal nValue As Variant)
    If IsObject(nValue) Then
        Set pValue = nValue
    Else
        Let pValue = nValue
    End If
    pType  = TypeName(nValue)
End Property

Public Property Get Value() As Variant
    If IsObject(pValue) Then
        Set Value = pValue
    Else
        Let Value = pValue
    End If
End Property

Public Property Get SourceType() As String
    SourceType = pType
End Property
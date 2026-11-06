VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryArgument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mArgumentReturn As QueryReturnType
Private mLink           As QueryArgumentLink
Private mValue          As Variant
Private mSource         As QueryDefinition


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal nArgumentReturn As QueryReturnType) As QueryArgument
    Set Create = New QueryArgument
    Call Create.Initialize(nArgumentReturn)
End Function

Friend Sub Initialize(ByVal nArgumentReturn As QueryReturnType)
    Set mArgumentReturn = nArgumentReturn
End Sub


Public Property Get ArgumentReturn() As QueryReturnType
    Set ArgumentReturn = mArgumentReturn
End Property


Public Property Get Link() As QueryArgumentLink
    Set Link = mLink
End Property

Public Property Set Link(ByVal nLink As QueryArgumentLink)
    Set mLink = nLink
End Property


Public Property Get Value() As Variant
    Value = mValue
End Property

Public Property Let Value(ByVal nValue As Variant)
    Let mValue = nValue
End Property


Public Property Get Source() As QueryDefinition
    Set Source = mSource
End Property

Public Property Set Source(ByVal nSource As QueryDefinition)
    Set mSource = nSource
End Property
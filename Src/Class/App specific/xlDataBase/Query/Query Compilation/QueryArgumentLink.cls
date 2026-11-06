VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryArgumentLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mInputIndex  As Long
Private mOutputIndex As Long
Private mIsConstant  As Boolean

'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal nInputIndex As Long, ByVal nOutputIndex As Long) As QueryArgumentLink
    Set Create = New QueryArgumentLink
    Call Create.Initialize(nInputIndex, nOutputIndex)
End Function

Friend Sub Initialize(ByVal nInputIndex As Long, ByVal nOutputIndex As Long)
    Let mInputIndex  = nInputIndex
    Let mOutputIndex = nOutputIndex
    Let mIsConstant  = False
End Sub

'===============================================================================
' PROPERTIES
'===============================================================================

Public Property Get InputIndex() As Long
    InputIndex = mInputIndex
End Property

Public Property Get OutputIndex() As Long
    OutputIndex = mOutputIndex
End Property

Public Property Get IsConstant() As Boolean
    IsConstant = mIsConstant
End Property
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryReturnType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mName As String
Private mType As QueryValueKind


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal nName As String, ByVal nType As QueryValueKind) As QueryReturnType
    Set Create = New QueryReturnType
    Call Create.Initialize(nName, nType)
End Function

Friend Sub Initialize(ByVal nName As String, ByVal nType As QueryValueKind)
    Let mName  = nName
    Let mType  = nType
End Sub

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get ReturnType() As QueryValueKind
    Let ReturnType = mType
End Property
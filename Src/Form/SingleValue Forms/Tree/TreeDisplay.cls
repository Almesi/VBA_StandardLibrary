VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mDisplayName As String
Private mValue As Variant
Private mChildren As New Collection



Public Function Create() As TreeDisplay
    Set Create = New TreeDisplay
End Function

Public Property Let DisplayName(ByVal Value As String)    : Let mDisplayName = Value        : End Property
Public Property Let Value(ByVal Value       As Variant)   : Let mValue       = Value        : End Property
Public Property Set Value(ByVal Value       As Object)    : Set mValue       = Value        : End Property

Public Property Get DisplayName()           As String     : Let DisplayName  = mDisplayName : End Property
Public Property Get Children()              As Collection : Set Children     = mChildren    : End Property

Public Property Get Value() As Variant
    If IsObject(mValue) Then
        Set Value = mValue
    Else
        Let Value = mValue
    End If
End Property

Public Sub AddChild(ByVal Child As TreeDisplay)
    Call mChildren.Add(Child)
End Sub

Public Function HasChildren() As Boolean
    HasChildren = (mChildren.Count > 0)
End Function

Public Sub Assign(ByVal nValue As Variant)
    If ISObject(nValue) Then
        Set Value = nValue
    Else
        Let Value = nValue
    End If
End Sub

Public Sub Flatten(ByVal Result As Collection)
    Dim Child As TreeDisplay
    Call Result.Add(mValue)

    For Each Child In mChildren
        Call Child.Flatten(Result)
    Next Child
End Sub
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Implements IEnumerator

Private mItems As Collection
Private mIndex As Long

Public Sub Initialize(ByVal Items As Collection)
    If Items Is Nothing Then Err.Raise 91, "CollectionEnumerator", "Collection cannot be Nothing."
    Set mItems = Items
    mIndex = 0
End Sub

Private Function MoveNextCore() As Boolean
    If mItems Is Nothing Then Err.Raise 91, "CollectionEnumerator", "Call Initialize first."
    If mIndex < mItems.Count Then
        mIndex = mIndex + 1
        MoveNextCore = True
    Else
        MoveNextCore = False
    End If
End Function

Private Property Get CurrentCore() As Variant
    If Not HasCurrentCore Then Err.Raise 5, "CollectionEnumerator", "Enumerator has no current element."
    If IsObject(mItems(mIndex)) Then
        Set CurrentCore = mItems(mIndex)
    Else
        CurrentCore = mItems(mIndex)
    End If
End Property

Private Property Get HasCurrentCore() As Boolean
    If mItems Is Nothing Then
        HasCurrentCore = False
        Exit Property
    End If
    HasCurrentCore = (mIndex >= 1 And mIndex <= mItems.Count)
End Property

Private Sub ResetCore()
    mIndex = 0
End Sub

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNextCore()
End Function

Private Property Get IEnumerator_Current() As Variant
    If IsObject(CurrentCore) Then
        Set IEnumerator_Current = CurrentCore
    Else
        IEnumerator_Current = CurrentCore
    End If
End Property

Private Property Get IEnumerator_HasCurrent() As Boolean
    IEnumerator_HasCurrent = HasCurrentCore
End Property

Private Sub IEnumerator_Reset()
    ResetCore
End Sub
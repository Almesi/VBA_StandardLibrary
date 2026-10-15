VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DataTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mRowEntry As RowEntry
Private mParent   As DataSetTreeNode
Private mChildren As New Collection

Public Function Create(ByVal Value As RowEntry) As DataTreeNode
    Set Create = New DataTreeNode
    Call Create.Initialize(Value)
End Function

Friend Sub Initialize(ByVal Value As RowEntry)
    Set mRowEntry = Value
End Sub

Public Property Get RowEntry() As RowEntry
    Set RowEntry = mRowEntry
End Property

Public Property Get Parent() As DataSetTreeNode
    Set Parent = mParent
End Property

Public Property Get Children() As Collection
    Set Children = mChildren
End Property

Public Property Get HasChildren() As Boolean
    HasChildren = (mChildren.Count > 0)
End Property

Public Sub AddChild(ByVal Child As DataSetTreeNode)
    If Child Is Nothing Then Exit Sub

    Set Child.mParent = Me
    Call mChildren.Add(Child)
End Sub
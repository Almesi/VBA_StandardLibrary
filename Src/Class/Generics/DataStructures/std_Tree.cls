
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


' ==================================================
' ====================[Properties]==================
' ==================================================

Private mRootNode   As std_Node
Private mCount      As Long

Private mEnumerator As std_NodeEnumerator


' ==================================================
' ====================[Creation]====================
' ==================================================

Public Function Create() As std_Tree
    Set Create = New std_Tree
End Function

Public Function CreateWithValue(Optional ByVal Value As Variant) As std_Tree
    Set CreateWithValue = Create()
    Call CreateWithValue.SetRoot(Value)
End Function


' ==================================================
' ====================[Getters/Setters]=============
' ==================================================

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = (mCount = 0)
End Property

Public Property Get RootNode() As std_Node
    Set RootNode = mRootNode
End Property

Public Property Get NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4

    Set mEnumerator = std_NodeEnumerator.Create(mRootNode, True)
    Set NewEnum = mEnumerator.[_NewEnum]
End Property


' ==================================================
' ====================[Methods]=====================
' ==================================================

Public Function SetRoot(Optional ByVal Value As Variant) As std_Node
    If Not mRootNode Is Nothing Then
        Err.Raise vbObjectError + 1600, "std_Tree.SetRoot", "Tree already has a root node."
    End If

    Set mRootNode = std_Node.Create(Value)
    mCount = 1

    Set SetRoot = mRootNode
End Function

Public Function AddChild(ByVal ParentNode As std_Node, Optional ByVal Value As Variant) As std_Node
    If mRootNode Is Nothing Then
        Err.Raise vbObjectError + 1601, "std_Tree.AddChild", "Tree has no root node."
    End If

    If ParentNode Is Nothing Then
        Err.Raise vbObjectError + 1602, "std_Tree.AddChild", "Parent node cannot be Nothing."
    End If

    If Not ContainsNode(ParentNode) Then
        Err.Raise vbObjectError + 1603, "std_Tree.AddChild", "Parent node does not belong to this tree."
    End If

    Set AddChild = std_Node.Create(Value)

    Call ParentNode.AddReference(AddChild)

    mCount = mCount + 1
End Function

Public Function GetParent(ByVal Node As std_Node) As std_Node
    If mRootNode Is Nothing Then Exit Function
    If Node Is Nothing Then Exit Function
    If Node Is mRootNode Then Exit Function

    Set GetParent = mRootNode.GetParent(Node, True)
End Function

Public Function ContainsNode(ByVal Node As std_Node) As Boolean
    If mRootNode Is Nothing Then Exit Function
    If Node Is Nothing Then Exit Function

    ContainsNode = (Node Is mRootNode Or mRootNode.HasReference(Node, True))
End Function

Public Sub Clear()
    Dim CurrentNode As std_Node

    If mRootNode Is Nothing Then Exit Sub

    Set mEnumerator = std_NodeEnumerator.Create(True, mRootNode)

    For Each CurrentNode In mEnumerator
        Call CurrentNode.ClearReferences()
    Next CurrentNode

    Set mRootNode = Nothing
    mCount = 0
End Sub


' ==================================================
' ====================[Helpers]=====================
' ==================================================

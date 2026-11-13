
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "std_ParentTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const vfRootRequired   As Long = 1
Private Const vfNodeRequired   As Long = 2
Private Const vfNodeMember     As Long = 4
Private Const vfParentRequired As Long = 8
Private Const vfParentMember   As Long = 16
Private Const vfRootForbidden  As Long = 32
Private Const vfNodeDetached   As Long = 64

Private mRootNode   As std_Node
Private mNodes      As Collection
Private mParents    As Collection
Private mCount      As Long
Private mValidator  As IValidator
Private mEnumerator As std_NodeEnumerator


Public Function Create(Optional ByVal Value As Variant) As std_ParentTree
    Set Create = CreateTyped(Nothing, Value)
End Function

Public Function CreateTyped(ByVal Validator As IValidator, Optional ByVal Value As Variant) As std_ParentTree
    Set CreateTyped = New std_ParentTree
    Call CreateTyped.Initialize(Validator, Value)
End Function

Friend Sub Initialize(ByVal Validator As IValidator, ByVal Value As Variant)
    Set mValidator = Validator
    If Not IsMissing(Value) Then Call SetRoot(Value)
End Sub

Private Sub Class_Initialize()
    Set mNodes = New Collection
    Set mParents = New Collection
End Sub


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
    Set NewEnum = mEnumerator.NewEnum
End Property


Public Function SetRoot(Optional ByVal Value As Variant) As std_Node
    Call ValidateInput("std_ParentTree.SetRoot")

    Set mRootNode = std_Node.CreateTyped(mValidator, Value)
    Call mNodes.Add(mRootNode, GetNodeKey(mRootNode))
    mCount = 1
    Set SetRoot = mRootNode
End Function


Public Function AddChild(ByVal ParentNode As std_Node, Optional ByVal Value As Variant) As std_Node
    Call ValidateInput("std_ParentTree.AddChild", , ParentNode, (vfRootRequired Or vfParentRequired Or vfParentMember))

    Dim ChildNode As std_Node
    Set ChildNode = std_Node.CreateTyped(mValidator, Value)

    Set AddChild = AttachSubtree(ParentNode, ChildNode)
End Function


Public Function AttachSubtree(ByVal ParentNode As std_Node, ByVal ChildNode As std_Node) As std_Node
    Call ValidateInput("std_ParentTree.AttachSubtree", ChildNode, ParentNode, (vfRootRequired Or vfNodeRequired Or vfNodeDetached Or vfParentRequired Or vfParentMember))

    Dim Visited As New Collection
    Call ValidateDetachedSubtree(ChildNode, Visited)

    Call ParentNode.AddReference(ChildNode)
    Call RegisterSubtree(ChildNode, ParentNode)

    Set AttachSubtree = ChildNode
End Function


Public Function DetachSubtree(ByVal Node As std_Node) As std_Node
    Call ValidateInput("std_ParentTree.DetachSubtree", Node, , (vfRootRequired Or vfNodeRequired Or vfNodeMember Or vfRootForbidden))

    Dim ParentNode As std_Node
    Set ParentNode = GetParent(Node)

    Call ParentNode.RemoveReference(Node)
    Call UnregisterSubtree(Node)

    Set DetachSubtree = Node

End Function


Public Sub RemoveSubtree(ByVal Node As std_Node)
    Call ValidateInput("std_ParentTree.RemoveSubtree", Node, , (vfRootRequired Or vfNodeRequired Or vfNodeMember Or vfRootForbidden))

    Dim Nodes As New Collection
    Dim CurrentNode As std_Node
    Dim DetachedNode As std_Node

    Call CollectSubtree(Node, Nodes)
    Set DetachedNode = DetachSubtree(Node)

    For Each CurrentNode In Nodes
        Call CurrentNode.ClearReferences()
    Next CurrentNode
End Sub


Public Sub MoveSubtree(ByVal Node As std_Node, ByVal NewParent As std_Node)
    Dim OldParent As std_Node
    Dim Nodes As New Collection
    Dim CurrentNode As std_Node

    Call ValidateInput("std_ParentTree.MoveSubtree", Node, NewParent, (vfRootRequired Or vfNodeRequired Or vfNodeMember Or vfRootForbidden Or vfParentRequired Or vfParentMember))

    If Node Is NewParent Then
        Err.Raise vbObjectError + 1621, "std_ParentTree.MoveSubtree", "A node cannot be moved below itself."
    End If

    Set OldParent = GetParent(Node)

    If OldParent Is NewParent Then Exit Sub

    Call CollectSubtree(Node, Nodes)

    For Each CurrentNode In Nodes
        If CurrentNode Is NewParent Then
            Err.Raise vbObjectError + 1622, "std_ParentTree.MoveSubtree", "A node cannot be moved below one of its descendants."
        End If
    Next CurrentNode

    Call OldParent.RemoveReference(Node)
    Call UnregisterSubtree(Node)

    Call NewParent.AddReference(Node)
    Call RegisterSubtree(Node, NewParent)
End Sub


Public Sub Clear()
    Dim CurrentNode As std_Node

    For Each CurrentNode In mNodes
        Call CurrentNode.ClearReferences()
    Next CurrentNode

    Set mNodes = New Collection
    Set mParents = New Collection
    Set mRootNode = Nothing

    mCount = 0

End Sub


Public Function GetParent(ByVal Node As std_Node) As std_Node
    If Node Is Nothing Then Exit Function
    If Node Is mRootNode Then Exit Function
    If Not ContainsNode(Node) Then Exit Function

    If CollectionContainsKey(mParents, GetNodeKey(Node)) Then
        Set GetParent = mParents(GetNodeKey(Node))
    End If
End Function


Public Function GetRoot(ByVal Node As std_Node) As std_Node
    Dim CurrentNode As std_Node
    Dim ParentNode As std_Node

    If Not ContainsNode(Node) Then Exit Function

    Set CurrentNode = Node
    Do
        Set ParentNode = GetParent(CurrentNode)
        If ParentNode Is Nothing Then Exit Do
        Set CurrentNode = ParentNode
    Loop

    Set GetRoot = CurrentNode

End Function


Public Function GetDepth(ByVal Node As std_Node) As Long
    Dim CurrentNode As std_Node
    Dim ParentNode As std_Node

    If Not ContainsNode(Node) Then Exit Function

    Set CurrentNode = Node

    Do
        Set ParentNode = GetParent(CurrentNode)
        If ParentNode Is Nothing Then Exit Do
        Set CurrentNode = ParentNode
        GetDepth = GetDepth + 1
    Loop
End Function


Public Function ContainsNode(ByVal Node As std_Node) As Boolean
    If Node Is Nothing Then Exit Function
    If Not CollectionContainsKey(mNodes, GetNodeKey(Node)) Then Exit Function

    ContainsNode = (mNodes(GetNodeKey(Node)) Is Node)
End Function


Public Function Validate() As Boolean
    Dim Visited As New Collection
    Dim CurrentNode As std_Node

    If mRootNode Is Nothing Then
        If mCount <> 0 Then Exit Function
        If mNodes.Count <> 0 Then Exit Function
        If mParents.Count <> 0 Then Exit Function
        Validate = True
        Exit Function
    End If

    If Not ContainsNode(mRootNode) Then Exit Function
    If CollectionContainsKey(mParents, GetNodeKey(mRootNode)) Then Exit Function

    If Not ValidateOwnedSubtree(mRootNode, Visited) Then Exit Function

    If Visited.Count  <> mCount Then Exit Function
    If mNodes.Count   <> mCount Then Exit Function
    If mParents.Count <> mCount - 1 Then Exit Function

    For Each CurrentNode In mNodes
        If Not CollectionContainsKey(Visited, GetNodeKey(CurrentNode)) Then Exit Function
    Next CurrentNode

    Validate = True

End Function


Private Sub ValidateInput(ByVal Source As String, Optional ByVal Node As std_Node, Optional ByVal ParentNode As std_Node, Optional ByVal Flags As Long = 0)

    If (Flags And vfRootRequired) <> 0 Then
        If mRootNode Is Nothing Then
            Err.Raise vbObjectError + 1600, Source, "Tree has no root node."
        End If
    End If

    If (Flags And vfNodeRequired) <> 0 Then
        If Node Is Nothing Then
            Err.Raise vbObjectError + 1601, Source, "Node cannot be Nothing."
        End If
    End If

    If (Flags And vfNodeMember) <> 0 Then
        If Not Node Is Nothing Then
            If Not ContainsNode(Node) Then
                Err.Raise vbObjectError + 1602, Source, "Node does not belong to this tree."
            End If
        End If
    End If

    If (Flags And vfParentRequired) <> 0 Then
        If ParentNode Is Nothing Then
            Err.Raise vbObjectError + 1603, Source, "Parent node cannot be Nothing."
        End If
    End If

    If (Flags And vfParentMember) <> 0 Then
        If Not ParentNode Is Nothing Then
            If Not ContainsNode(ParentNode) Then
                Err.Raise vbObjectError + 1604, Source, "Parent node does not belong to this tree."
            End If
        End If
    End If

    If (Flags And vfRootForbidden) <> 0 Then
        If Not Node Is Nothing Then
            If Node Is mRootNode Then
                Err.Raise vbObjectError + 1605, Source, "The root node cannot be used for this operation."
            End If
        End If
    End If

    If (Flags And vfNodeDetached) <> 0 Then
        If Not Node Is Nothing Then
            If ContainsNode(Node) Then
                Err.Raise vbObjectError + 1606, Source, "Node already belongs to this tree."
            End If
        End If
    End If

End Sub


Private Sub RegisterSubtree(ByVal Node As std_Node, ByVal ParentNode As std_Node)
    Dim ChildNode As std_Node
    Dim i As Long

    Call mNodes.Add(Node, GetNodeKey(Node))

    If Not ParentNode Is Nothing Then
        Call mParents.Add(ParentNode, GetNodeKey(Node))
    End If

    mCount = mCount + 1

    For i = 1 To Node.ReferenceCount
        Set ChildNode = Node.GetReference(i)
        Call RegisterSubtree(ChildNode, Node)
    Next i
End Sub


Private Sub UnregisterSubtree(ByVal Node As std_Node)
    Dim ChildNode As std_Node
    Dim i As Long

    For i = 1 To Node.ReferenceCount
        Set ChildNode = Node.GetReference(i)
        Call UnregisterSubtree(ChildNode)
    Next i

    Call RemoveCollectionKey(mParents, GetNodeKey(Node))
    Call RemoveCollectionKey(mNodes, GetNodeKey(Node))

    mCount = mCount - 1
End Sub


Private Sub ValidateDetachedSubtree(ByVal Node As std_Node, ByVal Visited As Collection)

    Dim ChildNode As std_Node
    Dim Key As String
    Dim i As Long

    If Node Is Nothing Then
        Err.Raise vbObjectError + 1623, "std_ParentTree.ValidateDetachedSubtree", "Subtree contains a Nothing reference."
    End If

    Key = GetNodeKey(Node)

    If CollectionContainsKey(Visited, Key) Then
        Err.Raise vbObjectError + 1624, "std_ParentTree.ValidateDetachedSubtree", "Subtree contains a cycle or duplicate node reference."
    End If

    If ContainsNode(Node) Then
        Err.Raise vbObjectError + 1625, "std_ParentTree.ValidateDetachedSubtree", "Subtree contains a node already belonging to this tree."
    End If

    Call Visited.Add(Node, Key)

    For i = 1 To Node.ReferenceCount
        Set ChildNode = Node.GetReference(i)
        Call ValidateDetachedSubtree(ChildNode, Visited)
    Next i

End Sub


Private Function ValidateOwnedSubtree(ByVal Node As std_Node, ByVal Visited As Collection) As Boolean
    Dim ChildNode As std_Node
    Dim Key As String
    Dim i As Long

    If Node Is Nothing Then Exit Function

    Key = GetNodeKey(Node)

    If CollectionContainsKey(Visited, Key) Then Exit Function
    If Not ContainsNode(Node) Then Exit Function

    Call Visited.Add(Node, Key)

    For i = 1 To Node.ReferenceCount

        Set ChildNode = Node.GetReference(i)

        If ChildNode Is Nothing Then Exit Function
        If Not ContainsNode(ChildNode) Then Exit Function

        If Not CollectionContainsKey(mParents, GetNodeKey(ChildNode)) Then Exit Function
        If Not (mParents(GetNodeKey(ChildNode)) Is Node) Then Exit Function

        If Not ValidateOwnedSubtree(ChildNode, Visited) Then Exit Function

    Next i

    ValidateOwnedSubtree = True
End Function


Private Sub CollectSubtree(ByVal Node As std_Node, ByVal Result As Collection)
    Dim ChildNode As std_Node
    Dim i As Long

    Call Result.Add(Node)

    For i = 1 To Node.ReferenceCount
        Set ChildNode = Node.GetReference(i)
        Call CollectSubtree(ChildNode, Result)
    Next i
End Sub


Private Function GetNodeKey(ByVal Node As std_Node) As String
    GetNodeKey = CStr(Node.ID)
End Function


Private Function CollectionContainsKey(ByVal Items As Collection, ByVal Key As String) As Boolean
    Dim Item As Object
    On Error Resume Next
    Set Item = Items(Key)
    CollectionContainsKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


Private Sub RemoveCollectionKey(ByVal Items As Collection, ByVal Key As String)
    If CollectionContainsKey(Items, Key) Then
        Items.Remove Key
    End If
End Sub
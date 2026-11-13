Attribute VB_Name = "Test_std_ParentTree"
Option Explicit

Public Sub Run_std_ParentTree_Tests()
    Dim T As std_ParentTree
    Dim Root As std_Node
    Dim A As std_Node
    Dim B As std_Node
    Dim C As std_Node
    Dim D As std_Node
    Dim Detached As std_Node
    Dim DetachedChild As std_Node
    Dim Outside As std_Node
    Dim N As std_Node
    Dim Seen As Long

    Debug.Print "std_ParentTree:"
    Set T = std_ParentTree.Create(1)
    Set Root = T.RootNode

    Call AssertEqual(1, T.Count, "Create count")
    Call AssertTrue(T.Validate, "Initial Validate")
    Call AssertTrue(T.ContainsNode(Root), "Contains root")
    Call AssertObjectIsNothing(T.GetParent(Root), "Root parent")

    Set A = T.AddChild(Root, 2)
    Set B = T.AddChild(Root, 3)
    Set C = T.AddChild(A, 4)
    Set D = T.AddChild(C, 5)

    Call AssertEqual(5, T.Count, "AddChild count")
    Call AssertTrue(T.Validate, "Validate after AddChild")
    Call AssertTrue(T.GetParent(A) Is Root, "Parent A")
    Call AssertTrue(T.GetParent(C) Is A, "Parent C")
    Call AssertTrue(T.GetParent(D) Is C, "Parent D")
    Call AssertTrue(T.GetRoot(D) Is Root, "GetRoot descendant")
    Call AssertEqual(3, T.GetDepth(D), "GetDepth descendant")
    Call AssertEqual(0, T.GetDepth(Root), "GetDepth root")

    For Each N In T
        Seen = Seen + 1
    Next N
    Call AssertEqual(T.Count, Seen, "Enumeration count")

    Call T.MoveSubtree(C, B)
    Call AssertTrue(T.Validate, "Validate after move")
    Call AssertTrue(T.GetParent(C) Is B, "Moved subtree parent")
    Call AssertTrue(T.GetParent(D) Is C, "Moved descendant parent")
    Call AssertEqual(3, T.GetDepth(D), "Moved descendant depth")

    Set Detached = std_Node.Create(20)
    Set DetachedChild = std_Node.Create(21)
    Call Detached.AddReference(DetachedChild)
    Call T.AttachSubtree(A, Detached)

    Call AssertTrue(T.Validate, "Validate after attach subtree")
    Call AssertEqual(7, T.Count, "Attach subtree count")
    Call AssertTrue(T.GetParent(Detached) Is A, "Attached root parent")
    Call AssertTrue(T.GetParent(DetachedChild) Is Detached, "Attached child parent")

    Set N = T.DetachSubtree(Detached)
    Call AssertTrue(N Is Detached, "Detach returns subtree root")
    Call AssertTrue(T.Validate, "Validate after detach")
    Call AssertEqual(5, T.Count, "Detach subtree count")
    Call AssertTrue(Not T.ContainsNode(Detached), "Detached root not contained")
    Call AssertTrue(Not T.ContainsNode(DetachedChild), "Detached descendant not contained")
    Call AssertTrue(Detached.FirstReference Is DetachedChild, "Detached subtree remains connected")

    Call T.RemoveSubtree(C)
    Call AssertTrue(T.Validate, "Validate after remove")
    Call AssertEqual(3, T.Count, "Remove subtree count")
    Call AssertTrue(Not T.ContainsNode(C), "Removed root not contained")
    Call AssertTrue(Not T.ContainsNode(D), "Removed descendant not contained")

    Set Outside = std_Node.Create(30)
    Call AssertTrue(Not T.ContainsNode(Outside), "External node not contained")
    Call AssertObjectIsNothing(T.GetParent(Outside), "External parent lookup")

    Call T.Clear
    Call AssertEqual(0, T.Count, "Clear count")
    Call AssertTrue(T.IsEmpty, "Clear IsEmpty")
    Call AssertObjectIsNothing(T.RootNode, "Clear root")
    Call AssertTrue(T.Validate, "Validate empty tree")

    Set Root = T.SetRoot(40)
    Set A = T.AddChild(Root, 41)
    Call AssertTrue(T.Validate, "Reuse after clear")
    Call AssertEqual(2, T.Count, "Reuse count")

End Sub
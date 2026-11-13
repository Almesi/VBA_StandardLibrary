Attribute VB_Name = "Test_std_Tree"

Option Explicit

Public Sub Run_std_Tree_Tests()
    Dim T As std_Tree
    Dim Root As std_Node
    Dim A As std_Node
    Dim B As std_Node
    Dim C As std_Node
    Dim Outside As std_Node
    Dim N As std_Node
    Dim Seen As Long

    Debug.Print "std_Tree:"
    Set T = std_Tree.Create(1)
    Set Root = T.RootNode
    Call AssertEqual(1, T.Count, "Create count")
    Call AssertTrue(Not T.IsEmpty, "Create IsEmpty")
    Call AssertTrue(T.ContainsNode(Root), "Contains root")

    Set A = T.AddChild(Root, 2)
    Set B = T.AddChild(Root, 3)
    Set C = T.AddChild(A, 4)

    Call AssertEqual(4, T.Count, "AddChild count")
    Call AssertTrue(T.ContainsNode(A), "Contains child")
    Call AssertTrue(T.ContainsNode(C), "Contains grandchild")
    Call AssertTrue(Not T.ContainsNode(Nothing), "Nothing is not contained")
    Call AssertTrue(T.GetParent(Root) Is Nothing, "Root has no parent")
    Call AssertTrue(T.GetParent(A) Is Root, "Parent of A")
    Call AssertTrue(T.GetParent(C) Is A, "Parent of C")

    Set Outside = std_Node.Create(99)
    Call AssertTrue(Not T.ContainsNode(Outside), "External node not contained")
    Call AssertTrue(T.GetParent(Outside) Is Nothing, "External node has no tree parent")

    For Each N In T
        Seen = Seen + 1
    Next N
    Call AssertEqual(T.Count, Seen, "Tree enumeration count")

    Call T.Clear
    Call AssertEqual(0, T.Count, "Clear count")
    Call AssertTrue(T.IsEmpty, "Clear IsEmpty")
    Call AssertObjectIsNothing(T.RootNode, "Clear root")

    Set Root = T.SetRoot(10)
    Set A = T.AddChild(Root, 11)
    Call AssertEqual(2, T.Count, "Reuse after Clear")
    Call AssertTrue(T.ContainsNode(A), "Contains child after reset")
End Sub
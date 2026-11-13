Attribute VB_Name = "Test_std_LinkedList"
Option Explicit

Public Sub Run_std_LinkedList_Tests()
    Dim L As std_LinkedList
    Dim N As std_Node
    Dim A As Variant
    Dim Item As Variant
    Dim i As Long
    Dim Expected As Variant

    Debug.Print "std_LinkedList:"
    Set L = std_LinkedList.Create(2)
    Call AssertEqual(1, L.Count, "Create count")
    Call AssertEqual(2, L.FirstNode.Value, "Create first value")
    Call AssertEqual(2, L.LastNode.Value, "Create last value")

    Call L.Push(4)
    Call L.Push(6)
    Call AssertEqual(3, L.Count, "Push count")
    Call AssertEqual(6, L.Peek.Value, "Peek returns last node")

    Call L.Insert(1, 1)
    Call L.Insert(3, 3)
    Call L.Insert(5, 5)
    Call AssertEqual(6, L.Count, "Insert count")

    Expected = Array(1, 2, 3, 4, 5, 6)
    A = L.ToArray
    For i = LBound(Expected) To UBound(Expected)
        Call AssertEqual(Expected(i), A(i), "Inserted order " & CStr(i))
    Next i

    Set N = L.GetNode(3)
    Call AssertEqual(3, N.Value, "GetNode by index")
    Call AssertEqual(3, L.IndexOf(N), "IndexOf")
    Call AssertTrue( L.ContainsNode(N), "ContainsNode true")

    Set N = L.GetNode()
    Call AssertEqual(6, N.Value, "GetNode default returns last")

    Set N = L.Remove(3)
    Call AssertEqual(3, N.Value, "Remove returns removed node")
    Call AssertEqual(5, L.Count, "Remove count")
    Call AssertEqual(4, L.GetNode(3).Value, "Remove reconnects list")

    Set N = L.Pop()
    Call AssertEqual(6, N.Value, "Pop returns last")
    Call AssertEqual(5, L.LastNode.Value, "Pop updates last node")

    Call L.Reverse
    Expected = Array(5, 4, 2, 1)
    A = L.ToArray
    For i = LBound(Expected) To UBound(Expected)
        Call AssertEqual(Expected(i), A(i), "Reverse order " & CStr(i))
    Next i

    Call L.Clear
    Call AssertEqual(0, L.Count, "Clear count")
    Call AssertTrue( L.IsEmpty, "Clear IsEmpty")
    Call AssertObjectIsNothing(L.FirstNode, "Clear first node")
    Call AssertObjectIsNothing(L.LastNode, "Clear last node")

    Call L.Push(10)
    Call AssertEqual(1, L.Count, "Push after Clear")
    Call AssertEqual(10, L.Peek.Value, "Peek after Clear")

    Call AssertRaises(1202, "LinkedList Pop empty", 1)
    Call AssertRaises(1208, "LinkedList Peek empty", 2)
    Call AssertRaises(1204, "LinkedList Insert invalid index", 3)
    Call AssertRaises(1205, "LinkedList Remove invalid index", 4)
    Call AssertRaises(1207, "LinkedList GetNode invalid index", 5)
End Sub

Private Sub AssertRaises(ByVal Number As Long, ByVal Message As String, ByVal TestCase As Long)
    Dim L As std_LinkedList
    Dim N As std_Node

    Set L = std_LinkedList.Create(1)
    If TestCase = 1 Or TestCase = 2 Then Call L.Clear

    On Error Resume Next
    Select Case TestCase
        Case 1
            Set N = L.Pop
        Case 2
            Set N = L.Peek
        Case 3
            Call L.Insert(3, 2)
        Case 4
            Set N = L.Remove(2)
        Case 5
            Set N = L.GetNode(2)
    End Select
    Dim ActualError As Long
    ActualError = Err.Number
    Err.Clear
    On Error GoTo 0
    If ActualError <> vbObjectError + Number Then Err.Raise vbObjectError + 9100, "Test_std_LinkedList", Message
End Sub
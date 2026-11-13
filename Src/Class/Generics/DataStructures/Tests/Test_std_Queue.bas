Attribute VB_Name = "Test_std_Queue"
Option Explicit

Public Sub Run_std_Queue_Tests()
    Dim Q As std_Queue
    Dim N As std_Node
    Dim Expected As Variant
    Dim Actual As Variant
    Dim i As Long

    Debug.Print "std_Queue:"
    Set Q = std_Queue.Create(1)
    Call AssertEqual(1, Q.Count, "Create count")
    Call AssertEqual(1, Q.Peek.Value, "Peek returns front")

    Call Q.Enqueue(2)
    Call Q.Enqueue(3)
    Call AssertEqual(3, Q.Count, "Enqueue count")
    Call AssertEqual(3, Q.BackNode.Value, "BackNode after enqueue")

    Expected = Array(1, 2, 3)
    Actual = QueueValues(Q)
    For i = 0 To 2
        Call AssertEqual(Expected(i), Actual(i), "Queue order " & CStr(i))
    Next i

    Set N = Q.Dequeue
    Call AssertEqual(1, N.Value, "Dequeue returns front")
    Call AssertEqual(2, Q.Peek.Value, "Dequeue advances front")
    Call AssertEqual(2, Q.Count, "Dequeue count")

    Set N = Q.Dequeue
    Call AssertEqual(2, N.Value, "Second dequeue")
    Set N = Q.Dequeue
    Call AssertEqual(3, N.Value, "Third dequeue")
    Call AssertTrue(Q.IsEmpty, "Queue empty after draining")
    Call AssertObjectIsNothing(Q.FrontNode, "FrontNode after draining")
    Call AssertObjectIsNothing(Q.BackNode, "BackNode after draining")

    Call Q.Enqueue(9)
    Call AssertEqual(9, Q.Peek.Value, "Enqueue after drain")
    Call Q.Clear
    Call AssertTrue(Q.IsEmpty, "Clear IsEmpty")
    Call AssertEqual(0, Q.Count, "Clear count")

    Call AssertRaises(1502, 1)
    Call AssertRaises(1503, 2)
End Sub

Private Function QueueValues(ByVal Q As std_Queue) As Variant
    Dim Result() As Variant
    Dim N As std_Node
    Dim i As Long

    ReDim Result(0 To Q.Count - 1)
    Set N = Q.FrontNode
    For i = 0 To Q.Count - 1
        Result(i) = N.Value
        Set N = N.FirstReference
    Next i
    QueueValues = Result
End Function

Private Sub AssertRaises(ByVal Number As Long, ByVal TestCase As Long)
    Dim Q As std_Queue
    Dim N As std_Node

    Set Q = std_Queue.Create(1)
    Call Q.Clear
    On Error Resume Next
    If TestCase = 1 Then Set N = Q.Dequeue
    If TestCase = 2 Then Set N = Q.Peek
    Dim ActualError As Long
    ActualError = Err.Number
    Err.Clear
    On Error GoTo 0
    If ActualError <> vbObjectError + Number Then Err.Raise vbObjectError + 9100, "Test_std_Queue", "Expected queue error " & CStr(Number)
End Sub
Attribute VB_Name = "Test_std_Stack"
Option Explicit

Public Sub Run_std_Stack_Tests()
    Dim S As std_Stack
    Dim N As std_Node

    Debug.Print "std_Stack:"
    Set S = std_Stack.Create(1)
    Call AssertEqual(1, S.Count, "Create count")
    Call AssertEqual(1, S.Peek.Value, "Peek returns top")

    Call S.Push(2)
    Call S.Push(3)
    Call AssertEqual(3, S.Count, "Push count")
    Call AssertEqual(3, S.TopNode.Value, "TopNode after push")

    Set N = S.Pop
    Call AssertEqual(3, N.Value, "Pop returns most recent")
    Call AssertEqual(2, S.Peek.Value, "Pop updates top")

    Set N = S.Pop
    Call AssertEqual(2, N.Value, "Second pop")
    Set N = S.Pop
    Call AssertEqual(1, N.Value, "Third pop")
    Call AssertTrue(S.IsEmpty, "Stack empty after draining")
    Call AssertObjectIsNothing(S.TopNode, "TopNode after draining")

    Call S.Push(7)
    Call AssertEqual(7, S.Peek.Value, "Push after drain")
    Call S.Clear
    Call AssertTrue(S.IsEmpty, "Clear IsEmpty")
    Call AssertEqual(0, S.Count, "Clear count")

    Call AssertRaises(1401, 1)
    Call AssertRaises(1402, 2)
End Sub

Private Sub AssertRaises(ByVal Number As Long, ByVal TestCase As Long)
    Dim S As std_Stack
    Dim N As std_Node

    Set S = std_Stack.Create(1)
    Call S.Clear
    On Error Resume Next
    If TestCase = 1 Then Set N = S.Pop
    If TestCase = 2 Then Set N = S.Peek
    Dim ActualError As Long
    ActualError = Err.Number
    Err.Clear
    On Error GoTo 0
    If ActualError <> vbObjectError + Number Then Err.Raise vbObjectError + 9100, "Test_std_Stack", "Expected stack error " & CStr(Number)
End Sub